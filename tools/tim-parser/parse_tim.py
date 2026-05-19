#!/usr/bin/env python3
"""
Parser de relatórios de dados cadastrais TIM.
Extrai registros de PDFs dentro de ZIPs para DataFrame, Excel e SQLite.
Inclui: deduplicação, snapshots históricos, estilização visual do Excel e geocodificação opcional.
"""

import argparse
import logging
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

# Adiciona tools/shared ao PYTHONPATH para importar modulos compartilhados
_shared_dir = Path(__file__).resolve().parent.parent / "shared"
sys.path.insert(0, str(_shared_dir))

import fitz  # PyMuPDF
import pandas as pd

import geocoding
from pipeline import deduplicar_dataframe, estilizar_excel, salvar_sqlite

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class RegistroTim:
    cpf_consulta: str
    numero_linha: str
    tipo_linha: str
    status_atual: str
    data_status: str
    data_inicio_vinculo: str
    data_cadastro: str
    data_fim_vinculo: str | None
    nome: str
    tipo_cliente: str
    cpf_cnpj: str
    sexo: str
    tipo_documento: str | None
    data_nascimento: str
    num_documento: str | None
    nacionalidade: str
    data_emissao: str | None
    telefone_contato: str | None
    pais_emissor: str | None
    endereco_residencial: str
    cidade_uf_cep_residencial: str
    endereco_fatura: str | None
    cidade_uf_cep_fatura: str | None
    arquivo_origem: str


# Regex para extrair CPF do header do relatório
RE_CPF_HEADER = re.compile(r"Relatório de Cadastro por \(CPF\):\s*(\d+)")
# Regex para extrair CPF do nome do arquivo: ..._CPF_...pdf
RE_CPF_ARQUIVO = re.compile(r"_(\d{11})_\d+\.pdf$")

# Linhas a serem ignoradas como lixo ou headers/footers
LIXO_EXATO = {"", "a", "aa", "CONFIDENCIAL", "TIM S/A", "DADOS CADASTRAIS",
              "EVENTOS DE PORTABILIDADE", "DADOS DE HABILITAÇÃO"}
RE_HEADER_FOOTER = re.compile(
    r"^(Número Solicitação:\s*\d+|"
    r"Número Protocolo:\s*.+|"
    r"Número Processo:\s*.+|"
    r"Rua Alexandre de Gusmão.*|"
    r"Santo André\s*[-–]\s*SP|"
    r"\+55\s*11\s*4251-6633|"
    r"\d+\s*/\s*\d+$|"
    r"Período de Pesquisa:.*|"
    r"Relatório de Cadastro por\s*\(CPF\):.*|"
    r"Não foram encontrados registros referentes à informação solicitada\.?)$"
)
RE_LABEL = re.compile(r"^([A-ZÀ-ÚÇÃÕÂÊÎÔÛ/\s\-]+):$")

# Labels que iniciam um novo registro (com dois-pontos)
LABEL_INICIO_REGISTRO = "NÚMERO DA LINHA:"

# Labels de campos válidos (para evitar capturar lixo como campo)
CAMPOS_VALIDOS = {
    "NÚMERO DA LINHA:", "TIPO DA LINHA:", "STATUS ATUAL:", "DATA STATUS:",
    "DATA INÍCIO VÍNCULO:", "DATA CADASTRO:", "DATA FIM VÍNCULO:",
    "NOME:", "TIPO DO CLIENTE:", "CPF/CNPJ:", "SEXO:", "TIPO DOCUMENTO:",
    "DATA NASCIMENTO:", "NÚM. DOCUMENTO:", "NACIONALIDADE:", "DATA EMISSÃO:",
    "TELEFONE CONTATO:", "PAÍS EMISSOR:", "ENDEREÇO RESIDENCIAL:",
    "CIDADE/UF - CEP RESIDENCIAL:", "ENDEREÇO FATURA:", "CIDADE/UF - CEP FATURA:",
}


def extrair_texto_pdf_de_bytes(data: bytes) -> str:
    """Abre PDF a partir de bytes e retorna texto concatenado de todas as páginas."""
    # Validacao de magic number para evitar processamento de arquivos nao-PDF
    if not data.startswith(b"%PDF"):
        logger.warning("Arquivo rejeitado: magic number invalido (nao e PDF)")
        return ""
    texto_paginas: list[str] = []
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        for page in doc:
            texto_paginas.append(page.get_text())
        doc.close()
    except Exception as e:
        logger.warning("Erro ao abrir PDF: %s", e)
    return "\n".join(texto_paginas)


def extrair_cpf_do_nome_arquivo(nome: str) -> str | None:
    """Tenta extrair CPF do nome do arquivo PDF."""
    m = RE_CPF_ARQUIVO.search(nome)
    if m:
        return m.group(1)
    return None


def extrair_cpf_do_texto(texto: str) -> str | None:
    """Tenta extrair CPF do header do relatório no texto."""
    m = RE_CPF_HEADER.search(texto)
    if m:
        return m.group(1)
    return None


def filtrar_linhas(linhas: list[str]) -> list[str]:
    """Remove headers, footers e lixo das linhas do PDF. Limita tamanho de linha para mitigar ReDoS."""
    filtradas: list[str] = []
    for linha in linhas:
        stripped = linha.strip()
        if len(stripped) > MAX_LINE_LENGTH:
            logger.warning("Linha truncada (%d chars) para mitigar ReDoS", len(stripped))
            stripped = stripped[:MAX_LINE_LENGTH]
        if stripped in LIXO_EXATO:
            continue
        if RE_HEADER_FOOTER.match(stripped):
            continue
        filtradas.append(stripped)
    return filtradas


def parse_texto_tim(texto: str, cpf_consulta: str, arquivo_origem: str) -> list[RegistroTim]:
    """Parseia texto concatenado de um ou mais PDFs da TIM."""
    linhas = texto.splitlines()
    linhas = filtrar_linhas(linhas)
    n = len(linhas)
    registros: list[RegistroTim] = []

    i = 0
    while i < n:
        # Procura início de registro
        if linhas[i] == LABEL_INICIO_REGISTRO:
            reg, i = _parse_registro_tim(linhas, i, cpf_consulta, arquivo_origem)
            if reg:
                registros.append(reg)
            continue
        i += 1

    return registros


def _parse_registro_tim(
    linhas: list[str], idx: int, cpf_consulta: str, arquivo_origem: str
) -> tuple[RegistroTim | None, int]:
    """Parseia um único registro a partir do label NÚMERO DA LINHA:."""
    n = len(linhas)
    i = idx + 1  # pula o próprio label de início
    campos: dict[str, str] = {}
    # O label de início foi detectado em parse_texto_tim; a próxima linha é seu valor
    label_atual: str | None = LABEL_INICIO_REGISTRO

    while i < n:
        linha = linhas[i]

        # Próximo registro
        if linha == LABEL_INICIO_REGISTRO:
            break

        # Detecta label
        m = RE_LABEL.match(linha)
        if m:
            label = m.group(0).strip()  # inclui o dois-pontos
            if label in CAMPOS_VALIDOS:
                label_atual = label
                i += 1
                continue
            else:
                label_atual = None
                i += 1
                continue

        # Se temos um label_atual, essa linha é o valor
        if label_atual is not None:
            valor = linha.strip()
            if valor == ".":
                valor = ""
            # Concatena se já existir valor para esse label (não deve acontecer normalmente)
            if label_atual in campos:
                if valor:
                    campos[label_atual] += " " + valor
            else:
                campos[label_atual] = valor
            label_atual = None
            i += 1
            continue

        i += 1

    # Validação mínima
    if "NÚMERO DA LINHA:" not in campos or not campos.get("NÚMERO DA LINHA:"):
        return None, i

    return RegistroTim(
        cpf_consulta=cpf_consulta,
        numero_linha=campos.get("NÚMERO DA LINHA:", ""),
        tipo_linha=campos.get("TIPO DA LINHA:", ""),
        status_atual=campos.get("STATUS ATUAL:", ""),
        data_status=campos.get("DATA STATUS:", ""),
        data_inicio_vinculo=campos.get("DATA INÍCIO VÍNCULO:", ""),
        data_cadastro=campos.get("DATA CADASTRO:", ""),
        data_fim_vinculo=campos.get("DATA FIM VÍNCULO:") or None,
        nome=campos.get("NOME:", ""),
        tipo_cliente=campos.get("TIPO DO CLIENTE:", ""),
        cpf_cnpj=campos.get("CPF/CNPJ:", ""),
        sexo=campos.get("SEXO:", ""),
        tipo_documento=campos.get("TIPO DOCUMENTO:") or None,
        data_nascimento=campos.get("DATA NASCIMENTO:", ""),
        num_documento=campos.get("NÚM. DOCUMENTO:") or None,
        nacionalidade=campos.get("NACIONALIDADE:", ""),
        data_emissao=campos.get("DATA EMISSÃO:") or None,
        telefone_contato=campos.get("TELEFONE CONTATO:") or None,
        pais_emissor=campos.get("PAÍS EMISSOR:") or None,
        endereco_residencial=campos.get("ENDEREÇO RESIDENCIAL:", ""),
        cidade_uf_cep_residencial=campos.get("CIDADE/UF - CEP RESIDENCIAL:", ""),
        endereco_fatura=campos.get("ENDEREÇO FATURA:") or None,
        cidade_uf_cep_fatura=campos.get("CIDADE/UF - CEP FATURA:") or None,
        arquivo_origem=arquivo_origem,
    ), i


# Limites de seguranca para processamento de ZIPs
MAX_ZIP_SIZE_BYTES = 500 * 1024 * 1024        # 500 MB
MAX_PDF_SIZE_BYTES = 100 * 1024 * 1024        # 100 MB
MAX_PDFS_PER_ZIP = 1000
MAX_COMPRESSION_RATIO = 100
MAX_LINE_LENGTH = 4096


def _validar_tamanho_zip(zf: zipfile.ZipFile) -> bool:
    """Verifica se o ZIP nao e um zip bomb e esta dentro dos limites."""
    total_uncompressed = 0
    total_compressed = 0
    for info in zf.infolist():
        total_uncompressed += info.file_size
        total_compressed += info.compress_size
        ratio = info.file_size / max(info.compress_size, 1)
        if ratio > MAX_COMPRESSION_RATIO and info.file_size > 1024 * 1024:
            logger.error(
                "ZIP rejeitado: ratio de compressao suspeito para %s (%.0fx)",
                info.filename, ratio,
            )
            return False
    if total_uncompressed > MAX_ZIP_SIZE_BYTES:
        logger.error("ZIP rejeitado: tamanho descomprimido total %d MB > limite %d MB",
                     total_uncompressed // (1024 * 1024), MAX_ZIP_SIZE_BYTES // (1024 * 1024))
        return False
    return True


def _sanitizar_nome_arquivo(nome: str) -> str:
    """Remove path traversal de nomes de arquivo para uso em logs."""
    return Path(nome).name


def processar_zip(caminho_zip: Path) -> list[RegistroTim]:
    """Processa um arquivo .zip contendo PDFs da TIM em subpastas."""
    todos_registros: list[RegistroTim] = []
    try:
        with zipfile.ZipFile(caminho_zip, "r") as zf:
            if not _validar_tamanho_zip(zf):
                return []

            pdfs = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
            if len(pdfs) > MAX_PDFS_PER_ZIP:
                logger.error("ZIP rejeitado: %d PDFs > limite %d", len(pdfs), MAX_PDFS_PER_ZIP)
                return []
            if not pdfs:
                logger.warning("Nenhum PDF encontrado em %s", caminho_zip.name)
                return []

            for nome_pdf in pdfs:
                try:
                    info = zf.getinfo(nome_pdf)
                    if info.file_size > MAX_PDF_SIZE_BYTES:
                        logger.error("PDF rejeitado: %s (%d MB) > limite %d MB",
                                     _sanitizar_nome_arquivo(nome_pdf),
                                     info.file_size // (1024 * 1024),
                                     MAX_PDF_SIZE_BYTES // (1024 * 1024))
                        continue

                    data = zf.read(nome_pdf)
                    texto = extrair_texto_pdf_de_bytes(data)
                    if not texto.strip():
                        continue

                    # Determina CPF da consulta
                    cpf = extrair_cpf_do_texto(texto) or extrair_cpf_do_nome_arquivo(nome_pdf) or ""
                    if not cpf:
                        logger.warning("CPF não identificado para %s", _sanitizar_nome_arquivo(nome_pdf))

                    regs = parse_texto_tim(texto, cpf, f"{caminho_zip.name}/{nome_pdf}")
                    if regs:
                        logger.debug("%s/%s → %d registros", caminho_zip.name, _sanitizar_nome_arquivo(nome_pdf), len(regs))
                    todos_registros.extend(regs)
                except Exception as e:
                    logger.error("Erro ao processar %s: %s", _sanitizar_nome_arquivo(nome_pdf), e)
    except Exception as e:
        logger.error("Erro ao abrir ZIP %s: %s", caminho_zip, e)

    return todos_registros


def selecionar_snapshots_recentes(df: pd.DataFrame) -> pd.DataFrame:
    """Mantém apenas o snapshot mais recente por numero_linha (máximo data_cadastro)."""
    df = df.copy()
    df["_data_dt"] = pd.to_datetime(
        df["data_cadastro"], format="%d/%m/%Y %H:%M:%S", errors="coerce"
    )
    idx = df.groupby("numero_linha")["_data_dt"].idxmax()
    df = df.loc[idx].reset_index(drop=True)
    df = df.drop(columns=["_data_dt"])
    return df


def _validar_caminho(caminho: Path, nome: str) -> None:
    """Valida que o caminho é absoluto e não tenta path traversal."""
    try:
        caminho_resolvido = caminho.resolve()
    except (OSError, RuntimeError) as e:
        raise ValueError(f"Caminho inválido para {nome}: {caminho}") from e
    if ".." in caminho.parts or not caminho_resolvido.is_absolute():
        raise ValueError(f"Caminho não permitido para {nome}: {caminho}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Parser de dados cadastrais TIM")
    parser.add_argument("--input", required=True, help="Pasta com arquivos .zip da TIM")
    parser.add_argument("--output", required=True, help="Pasta de saída")
    parser.add_argument("--no-geocode", action="store_false", dest="geocode", default=True,
                        help="Desativar geocodificação (padrão: ativado)")
    parser.add_argument("--historico", action="store_true", default=False,
                        help="Manter todos os snapshots históricos (padrão: apenas mais recente)")
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_dir = Path(args.output)
    _validar_caminho(input_dir, "input")
    _validar_caminho(output_dir, "output")
    output_dir.mkdir(parents=True, exist_ok=True)

    arquivos = sorted(input_dir.glob("*.zip"))
    if not arquivos:
        logger.error("Nenhum .zip encontrado em %s", input_dir)
        return

    logger.info("Processando %d arquivos ZIP...", len(arquivos))

    todos_registros: list[RegistroTim] = []
    for arq in arquivos:
        try:
            regs = processar_zip(arq)
            logger.info("%s → %d registros", arq.name, len(regs))
            todos_registros.extend(regs)
        except Exception as e:
            logger.error("Erro em %s: %s", arq.name, e, exc_info=True)

    if not todos_registros:
        logger.error("Nenhum registro extraído.")
        return

    df_raw = pd.DataFrame([asdict(r) for r in todos_registros])

    # Snapshots históricos
    if args.historico:
        logger.info("Modo histórico: mantendo todos os snapshots.")
        df_dedup = deduplicar_dataframe(df_raw)
    else:
        logger.info("Selecionando snapshot mais recente por número de linha...")
        df_dedup = selecionar_snapshots_recentes(df_raw)
        # Deduplica adicional após seleção de snapshot
        df_dedup = deduplicar_dataframe(df_dedup)

    # Geocodificação (opcional)
    db_path = output_dir / "dados_cadastrais.db"
    if args.geocode:
        logger.info("Iniciando geocodificação...")
        # Compõe endereço completo para geocodificação
        df_dedup["_endereco_completo"] = (
            df_dedup["endereco_residencial"].astype(str) + ", " +
            df_dedup["cidade_uf_cep_residencial"].astype(str)
        )
        # Cria colunas temporárias compatíveis com geocoding.py
        df_geo = df_dedup.copy()
        df_geo["endereco"] = df_geo["_endereco_completo"]
        # Extrai bairro, municipio, estado, cep do campo cidade_uf_cep_residencial
        # Formato esperado: "CIDADE/UF - CEP"
        def _extrair_cidade_uf_cep(valor: str) -> tuple[str, str, str, str]:
            bairro = ""
            municipio = ""
            estado = ""
            cep = ""
            if " - " in valor:
                parte1, parte2 = valor.rsplit(" - ", 1)
                cep = parte2.strip()
                if "/" in parte1:
                    municipio, estado = parte1.rsplit("/", 1)
                    municipio = municipio.strip()
                    estado = estado.strip()
                else:
                    municipio = parte1.strip()
            else:
                municipio = valor.strip()
            return bairro, municipio, estado, cep

        partes = df_geo["cidade_uf_cep_residencial"].apply(lambda x: _extrair_cidade_uf_cep(str(x)))
        df_geo["bairro"] = partes.apply(lambda x: x[0])
        df_geo["municipio"] = partes.apply(lambda x: x[1])
        df_geo["estado"] = partes.apply(lambda x: x[2])
        df_geo["cep"] = partes.apply(lambda x: x[3])

        geocoding.processar_geocoding(df_geo, str(db_path))

        # Transfere colunas geocodificadas de volta
        df_dedup["latitude"] = df_geo["latitude"]
        df_dedup["longitude"] = df_geo["longitude"]
        df_dedup["google_maps_url"] = df_geo["google_maps_url"]
        df_dedup = df_dedup.drop(columns=["_endereco_completo"])
    else:
        logger.info("Geocodificação desativada.")
        if "_endereco_completo" in df_dedup.columns:
            df_dedup = df_dedup.drop(columns=["_endereco_completo"])

    # Excel estilizado
    excel_path = output_dir / "dados_cadastrais_tim.xlsx"
    estilizar_excel(
        df_dedup,
        excel_path,
        col_situacao="status_atual",
        col_modalidade="tipo_linha",
    )

    # SQLite
    salvar_sqlite(df_raw, df_dedup, db_path, tabela_raw="linhas_tim_raw", tabela_dedup="linhas_tim")

    # Resumo
    logger.info("=" * 50)
    logger.info("RESUMO")
    logger.info("Total de registros (raw): %d", len(df_raw))
    logger.info("Total de registros (dedup/snapshots): %d", len(df_dedup))
    logger.info("CPFs únicos consultados: %d", df_dedup["cpf_consulta"].nunique())
    logger.info("Nomes únicos: %d", df_dedup["nome"].nunique())
    logger.info("Status: %s", df_dedup["status_atual"].value_counts().to_dict())
    logger.info("=" * 50)


if __name__ == "__main__":
    main()
