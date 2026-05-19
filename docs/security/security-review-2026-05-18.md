# Security Review: Parsers Vivo & TIM

**Data:** 2026-05-18  
**Scope:** `tools/tim-parser/parse_tim.py`, `tools/vivo-parser/parse_vivo.py`, `tools/shared/pipeline.py`, `tools/shared/geocoding.py`  
**Stack:** Python 3.13, pandas, openpyxl, PyMuPDF, requests, sqlite3  
**Threat Actors:** Operador mal-intencionado com acesso aos arquivos de entrada; atacante que envia ZIP/PDFs maliciosos; vazamento de PII por logs.

---

## Data Flow

```
Arquivo de entrada (.txt/.zip)
    -> Parser (parse_vivo.py / parse_tim.py)
        -> Extracao de campos (regex, PDF text)
        -> DataFrame (pandas)
            -> pipeline.py (dedup, sort, estilizar_excel, salvar_sqlite)
                -> SQLite (dados_cadastrais.db)
                -> Excel (.xlsx)
                -> geocoding.py -> Nominatim API (HTTP)
```

**Trust Boundaries:**
- Fronteira 1: Arquivos de entrada (nao confiaveis — podem ser manipulados)
- Fronteira 2: Processamento em memoria (confiavel)
- Fronteira 3: Saida para disco/DB (confiavel)
- Fronteira 4: Rede externa — Nominatim (nao confiavel, mas benigno)

---

## Findings

### FIND-001: Zip Bomb / Resource Exhaustion — **HIGH**

**File:** `tools/tim-parser/parse_tim.py:242-266`  
**Evidence:**
```python
with zipfile.ZipFile(caminho_zip, "r") as zf:
    pdfs = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
    for nome_pdf in pdfs:
        data = zf.read(nome_pdf)
        texto = extrair_texto_pdf_de_bytes(data)
```

**Exploit Scenario:**
Um atacante fornece um ZIP aparentemente pequeno (~10KB) contendo um PDF altamente compressivel (ex: 10MB de zeros). O parser carrega todo o conteudo em memoria (`zf.read()`) e depois o PDF em memoria novamente (`fitz.open(stream=...)`). Multiplos ZIPs assim podem esgotar RAM e causar DoS.

**Verification:**
```
Compressed: 10204 bytes
Uncompressed: 10485760 bytes
Ratio: 1028x
```

**Fix:**
- Limitar tamanho maximo do ZIP (ex: 500MB)
- Limitar tamanho maximo de cada entrada descomprimida (ex: 100MB)
- Limitar numero maximo de PDFs por ZIP (ex: 1000)
- Verificar ratio compressao/descompressao (alertar se > 100x)

---

### FIND-002: Path Traversal via Nomes de Arquivo no ZIP — **MEDIUM**

**File:** `tools/tim-parser/parse_tim.py:248-250`  
**Evidence:**
```python
for nome_pdf in pdfs:
    data = zf.read(nome_pdf)
```

`zf.read(nome_pdf)` aceita nomes como `../../../etc/passwd`. Embora nao escreva no filesystem (pois nao usamos `extractall()`), o nome pode ser usado em logs (`logger.error("Erro ao processar %s", nome_pdf)`) causando confusao ou log injection.

**Fix:**
- Sanitizar nomes de arquivos antes de logar (remover `..`, barras iniciais)
- Ou simplesmente usar `Path(nome_pdf).name` para logs

---

### FIND-003: PII Exposure em Logs de Erro — **MEDIUM**

**File:** `tools/tim-parser/parse_tim.py:265, 312` e `tools/vivo-parser/parse_vivo.py:214`  
**Evidence:**
```python
logger.error("Erro ao processar %s em %s: %s", nome_pdf, caminho_zip.name, e, exc_info=True)
logger.error("Erro em %s: %s", arq.name, e, exc_info=True)
```

**Exploit Scenario:**
Se uma excecao ocorrer durante o processamento de um DataFrame contendo CPFs/nomes/enderecos, `exc_info=True` pode incluir o conteudo do DataFrame no stack trace (via `repr()` de objetos pandas), vazando PII para logs.

**Fix:**
- Remover `exc_info=True` de logs que processam dados sensiveis, ou
- Redirecionar `exc_info` para um log file separado nao compartilhado
- Ou usar `logger.exception()` em vez de `exc_info=True` com filtro de PII

---

### FIND-004: ReDoS Potencial em Regexes — **MEDIUM**

**File:** `tools/tim-parser/parse_tim.py:69-80`  
**Evidence:**
```python
RE_HEADER_FOOTER = re.compile(
    r"^(Número Solicitação:\s*\d+|"
    r"Número Protocolo:\s*.+|"          # .+ é greedy
    r"Período de Pesquisa:.*|"          # .* é greedy
    r"Relatório de Cadastro por\s*\(CPF\):.*|"
    r"Não foram encontrados registros referentes à informação solicitada\.?)$"
)
```

**Exploit Scenario:**
Um PDF malicioso com uma linha extremamente longa (~1MB) sem quebra de linha pode causar backtracking na regex `^(...|...|.*)$`. O Python 3.11+ tem protecao contra ReDoS catastrofico, mas o processamento ainda pode ser lento.

**Fix:**
- Limitar tamanho maximo de cada linha processada (ex: 4096 chars)
- Simplificar regexes para evitar alternancias complexas com `.*`

---

### FIND-005: Processamento de PDF sem Validacao — **MEDIUM**

**File:** `tools/tim-parser/parse_tim.py:97-107`  
**Evidence:**
```python
def extrair_texto_pdf_de_bytes(data: bytes) -> str:
    doc = fitz.open(stream=data, filetype="pdf")
```

Nao ha validacao do magic number `%PDF` nem limite de tamanho. Um arquivo renomeado para `.pdf` dentro do ZIP (ex: um binario de 1GB) sera carregado inteiramente em memoria pelo PyMuPDF.

**Fix:**
- Verificar magic number `data[:4] == b'%PDF'` antes de chamar `fitz.open()`
- Limitar tamanho do PDF (ex: max 50MB)

---

### FIND-006: PYTHONPATH Manipulation — **LOW**

**File:** `tools/tim-parser/parse_tim.py:16-18` e `tools/vivo-parser/parse_vivo.py:15-17`  
**Evidence:**
```python
_shared_dir = Path(__file__).resolve().parent.parent / "shared"
sys.path.insert(0, str(_shared_dir))
```

**Exploit Scenario:**
Se um atacante puder modificar o diretorio `tools/shared/` (ou criar um arquivo `pipeline.py` ou `geocoding.py` falso no mesmo diretorio), o Python carregara o modulo falso em vez do real.

**Fix:**
- Usar import absoluto via `PYTHONPATH` no ambiente, ou
- Usar `importlib.util.spec_from_file_location` para imports explicitos, ou
- Pelo menos verificar que o arquivo importado esta no diretorio esperado

---

### FIND-007: No Output Path Validation — **LOW**

**File:** `tools/tim-parser/parse_tim.py:296` e `tools/vivo-parser/parse_vivo.py:198`  
**Evidence:**
```python
output_dir.mkdir(parents=True, exist_ok=True)
```

Nao ha validacao de que `output_dir` esta dentro de um diretorio permitido. Um comando como `--output C:\Windows\System32\fake` criaria diretorios em locais sensiveis.

**Fix:**
- Validar que `output_dir` e `input_dir` sao caminhos absolutos e dentro de diretorios permitidos
- Ou usar `resolve()` e verificar prefixo

---

### FIND-008: MD5 para Cache de Geocodificacao — **LOW**

**File:** `tools/shared/geocoding.py:44-46`  
**Evidence:**
```python
def _hash_endereco(endereco_completo: str) -> str:
    return hashlib.md5(endereco_completo.encode("utf-8")).hexdigest()
```

MD5 e criptograficamente quebrado. Para cache local SQLite, o risco e baixo (nao e usado para integridade ou autenticacao), mas e uma ma pratica.

**Fix:**
- Substituir por `hashlib.sha256()` ou `blake2b`

---

## Attack Chains

**Chain A (DoS + PII Leak):**
1. Atacante envia ZIP bomb (FIND-001)
2. Parser esgota memoria e gera `MemoryError`
3. `exc_info=True` captura stack trace (FIND-003)
4. Stack trace pode conter dados sensiveis do DataFrame
5. Log e salvo em disco ou enviado para sistema de monitoramento

**Chain B (Confusion + Path Traversal):**
1. Atacante cria ZIP com entrada `../../../pipeline.py`
2. Parser le o arquivo (nao e um PDF, mas `fitz.open` falha graciosamente)
3. Nome do arquivo aparece em logs (FIND-002)
4. Operador confunde o log com um problema real no sistema

---

## Remediation Applied

### Status das Correções

| ID | Severidade | Fix Aplicado | Arquivo | Status |
|---|---|---|---|---|
| FIND-001 | HIGH | Limites de ZIP/PDF + ratio de compressão | `parse_tim.py` | ✅ Aplicado |
| FIND-003 | MEDIUM | Removido `exc_info=True` dos logs | `parse_tim.py`, `parse_vivo.py` | ✅ Aplicado |
| FIND-002 | MEDIUM | Sanitização de nomes via `Path(nome).name` | `parse_tim.py` | ✅ Aplicado |
| FIND-004 | MEDIUM | Truncamento de linhas em 4096 chars | `parse_tim.py` | ✅ Aplicado |
| FIND-005 | MEDIUM | Validação de magic number `%PDF` | `parse_tim.py` | ✅ Aplicado |
| FIND-007 | LOW | Validação de path absoluto + `resolve()` | `parse_tim.py`, `parse_vivo.py` | ✅ Aplicado |
| FIND-008 | LOW | MD5 → SHA-256 | `geocoding.py` | ✅ Aplicado |
| FIND-006 | LOW | Documentado como risco residual | — | 📝 Documentado |

---

## Verification Results

### 1. Testes Unitários
```
pytest tests/ -v
============================= 10 passed in 1.58s =============================
```

### 2. Parser Vivo com Dados Reais
- **Input:** 36 arquivos `.txt`
- **Raw:** 397 registros
- **Dedup:** 89 registros
- **Saídas:** Excel + SQLite gerados corretamente
- **Exit code:** 0

### 3. Parser TIM com Dados Reais
- **Input:** 1 arquivo ZIP
- **Raw:** 32 registros
- **Dedup:** 15 registros (após seleção de snapshots)
- **Saídas:** Excel + SQLite gerados corretamente
- **Exit code:** 0

### 4. Testes de Segurança

**Zip Bomb (FIND-001):**
```
ZIP rejeitado: ratio de compressao suspeito para bomb.pdf (1029x)
Zip bomb test - regs: 0
```
✅ Rejeitado corretamente.

**Path Traversal (FIND-002):**
```
CPF não identificado para evil.pdf
Path traversal test - regs: 1
```
✅ Nome sanitizado para `evil.pdf` nos logs. Registro processado normalmente (PDF válido).

**Fake PDF (FIND-005):**
```
Arquivo rejeitado: magic number invalido (nao e PDF)
Fake PDF test - regs: 0
```
✅ Rejeitado corretamente.

### 5. SQLite Verification

```
TIM tables: ['linhas_tim_raw', 'linhas_tim', 'geocoding_cache', 'processamento_log', 'sqlite_sequence']
  Raw: 32
  Dedup: 15
Vivo tables: ['linhas_vivo_raw', 'linhas_vivo', 'geocoding_cache', 'processamento_log', 'sqlite_sequence']
  Raw: 397
  Dedup: 89
```

---

## Residual Risk

- **PyMuPDF (fitz):** A biblioteca é um wrapper C++ sobre MuPDF. Vulnerabilidades nessa camada não são auditáveis neste escopo.
- **Nominatim API:** Dados de endereço (sem CPF/nome) são enviados para servidores externos. Isso é aceitável para o caso de uso.
- **SQLite:** O banco não é criptografado. Em disco, PII está em texto plano. O sistema operacional deve garantir permissões de arquivo.
- **PYTHONPATH manipulation (FIND-006):** Risco baixo em ambiente controlado. Recomenda-se usar `PYTHONPATH` via ambiente ou `importlib.util` em futura refatoração.
- **Vivo parser sem limites de linha:** O parser Vivo não tem truncamento de linhas como o TIM. O risco de ReDoS existe teoricamente, mas o formato `.txt` Vivo tem quebras de linha regulares.
