#!/usr/bin/env python3
"""
Parser de relatórios de dados cadastrais Vivo.
Extrai registros de arquivos .txt para DataFrame, Excel e SQLite.
Inclui: deduplicação, estilização visual do Excel e geocodificação opcional.
"""

import argparse
import logging
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

# Adiciona tools/shared ao PYTHONPATH para importar modulos compartilhados
_shared_dir = Path(__file__).resolve().parent.parent / "shared"
sys.path.insert(0, str(_shared_dir))

import pandas as pd

import geocoding
from pipeline import deduplicar_dataframe, estilizar_excel, ordenar_dataframe, salvar_sqlite

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


@dataclass
class RegistroVivo:
    cpf_consulta: str
    numero_linha: str
    cliente: str
    cpf: str
    endereco: str
    bairro: str
    cep: str
    municipio: str
    estado: str
    modalidade: str
    situacao: str
    data_habilitacao: str
    data_rescisao: str | None
    arquivo_origem: str


# Regex mais flexível para extrair label:valor com pontos no meio
RE_CAMPO = re.compile(r"^\*\s+([A-ZÀ-ÚÇÃÕÂÊÎÔÛ\s/]+):\s*\.*(.+?)$")
RE_CPF_CONSULTA = re.compile(r"^\*\s*CPF:\s+([\d\.\-]+)\s*\*$")
RE_NUMERO_LINHA = re.compile(r"^\*\s*NÚMERO DA LINHA:")
RE_BLOCO_VAZIO = re.compile(r"Nenhum Dado Foi Encontrado", re.IGNORECASE)


def parse_arquivo(caminho: Path) -> list[RegistroVivo]:
    """Parseia um único arquivo .txt e retorna lista de registros."""
    texto = caminho.read_text(encoding="utf-8")
    linhas = texto.splitlines()
    n = len(linhas)
    registros: list[RegistroVivo] = []

    i = 0
    cpf_consulta = ""

    while i < n:
        linha = linhas[i]

        # Detecta CPF de consulta
        m_cpf = RE_CPF_CONSULTA.match(linha)
        if m_cpf:
            cpf_consulta = limpar_valor(m_cpf.group(1))
            i += 1
            continue

        # Detecta início de registro
        if RE_NUMERO_LINHA.match(linha):
            reg, i = parse_registro(linhas, i, cpf_consulta, caminho.name)
            if reg:
                registros.append(reg)
            continue

        i += 1

    return registros


def limpar_valor(valor: str) -> str:
    """Remove trailing ' *' e espaços extras do valor extraído."""
    return valor.rstrip(" *").strip()


def parse_registro(
    linhas: list[str], idx: int, cpf_consulta: str, arquivo: str
) -> tuple[RegistroVivo | None, int]:
    """Parseia um único registro a partir da linha do NÚMERO DA LINHA.
    Retorna (registro, próximo_idx).
    """
    n = len(linhas)
    campos: dict[str, str] = {}
    i = idx

    while i < n:
        linha = linhas[i]

        # Próximo registro ou fim de bloco de relatório (novo CPF ou separador grosso)
        if i > idx and RE_NUMERO_LINHA.match(linha):
            break
        if i > idx and RE_CPF_CONSULTA.match(linha):
            break
        if i > idx and "PARÂMETRO(S) DE CONSULTA" in linha:
            break
        if i > idx and "********************************************************************************" in linha:
            break

        # Tenta extrair campo
        m = RE_CAMPO.match(linha)
        if m:
            label = m.group(1).strip()
            valor = limpar_valor(m.group(2))

            if label == "ENDEREÇO":
                # Endereço pode ter continuação na(s) próxima(s) linha(s)
                # se não forem campos conhecidos nem separadores vazios
                endereco = valor
                j = i + 1
                while j < n:
                    prox = linhas[j]
                    # Se for próximo campo conhecido ou separador, para
                    if RE_NUMERO_LINHA.match(prox) or RE_CPF_CONSULTA.match(prox):
                        break
                    if "PARÂMETRO(S)" in prox or "***" in prox:
                        break
                    m2 = RE_CAMPO.match(prox)
                    if m2 and m2.group(1).strip() in {
                        "CLIENTE", "CPF", "BAIRRO", "CEP", "MUNICÍPIO",
                        "ESTADO", "MODALIDADE", "SITUAÇÃO", "DATA HABILITAÇÃO", "DATA RESCISÃO",
                    }:
                        break
                    # Se não casou nada acima, é continuação do endereço
                    # (pode ser linha sem label, só texto)
                    if prox.strip() and not prox.strip().startswith("*"):
                        endereco += " " + prox.strip()
                        j += 1
                        continue
                    if prox.strip() == "":
                        j += 1
                        continue
                    # Linha com * mas não é campo conhecido → continuação
                    if prox.strip().startswith("*") and not m2:
                        # remove o * inicial e adiciona
                        resto = prox.strip().lstrip("*").strip()
                        if resto:
                            endereco += " " + resto
                        j += 1
                        continue
                    break
                campos["ENDEREÇO"] = endereco.strip()
                i = j
                continue
            else:
                campos[label] = valor

        i += 1

    # Validação mínima
    if "NÚMERO DA LINHA" not in campos or "CLIENTE" not in campos:
        return None, i

    return RegistroVivo(
        cpf_consulta=cpf_consulta,
        numero_linha=campos.get("NÚMERO DA LINHA", ""),
        cliente=campos.get("CLIENTE", ""),
        cpf=campos.get("CPF", ""),
        endereco=campos.get("ENDEREÇO", ""),
        bairro=campos.get("BAIRRO", ""),
        cep=campos.get("CEP", ""),
        municipio=campos.get("MUNICÍPIO", ""),
        estado=campos.get("ESTADO", ""),
        modalidade=campos.get("MODALIDADE", ""),
        situacao=campos.get("SITUAÇÃO", ""),
        data_habilitacao=campos.get("DATA HABILITAÇÃO", ""),
        data_rescisao=campos.get("DATA RESCISÃO") or None,
        arquivo_origem=arquivo,
    ), i


def main() -> None:
    parser = argparse.ArgumentParser(description="Parser de dados cadastrais Vivo")
    parser.add_argument("--input", required=True, help="Pasta com arquivos .txt")
    parser.add_argument("--output", required=True, help="Pasta de saída")
    parser.add_argument("--no-geocode", action="store_false", dest="geocode", default=True,
                        help="Desativar geocodificação (padrão: ativado)")
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    arquivos = sorted(input_dir.glob("*.txt"))
    if not arquivos:
        logger.error("Nenhum .txt encontrado em %s", input_dir)
        return

    logger.info("Processando %d arquivos...", len(arquivos))

    todos_registros: list[RegistroVivo] = []
    for arq in arquivos:
        try:
            regs = parse_arquivo(arq)
            logger.info("%s → %d registros", arq.name, len(regs))
            todos_registros.extend(regs)
        except Exception as e:
            logger.error("Erro em %s: %s", arq.name, e, exc_info=True)

    if not todos_registros:
        logger.error("Nenhum registro extraído.")
        return

    df_raw = pd.DataFrame([asdict(r) for r in todos_registros])

    # Deduplicação
    df_dedup = deduplicar_dataframe(df_raw)

    # Ordenação
    df_dedup = ordenar_dataframe(df_dedup)

    # Geocodificação (opcional)
    db_path = output_dir / "dados_cadastrais.db"
    if args.geocode:
        logger.info("Iniciando geocodificação...")
        geocoding.processar_geocoding(df_dedup, str(db_path))
    else:
        logger.info("Geocodificação desativada.")

    # Excel estilizado
    excel_path = output_dir / "dados_cadastrais_vivo.xlsx"
    estilizar_excel(df_dedup, excel_path)

    # SQLite
    salvar_sqlite(df_raw, df_dedup, db_path, tabela_raw="linhas_vivo_raw", tabela_dedup="linhas_vivo")

    # Resumo
    logger.info("=" * 50)
    logger.info("RESUMO")
    logger.info("Total de registros (raw): %d", len(df_raw))
    logger.info("Total de registros (dedup): %d", len(df_dedup))
    logger.info("CPFs únicos consultados: %d", df_dedup["cpf_consulta"].nunique())
    logger.info("Clientes únicos: %d", df_dedup["cliente"].nunique())
    logger.info("Situações: %s", df_dedup["situacao"].value_counts().to_dict())
    logger.info("=" * 50)


if __name__ == "__main__":
    main()
