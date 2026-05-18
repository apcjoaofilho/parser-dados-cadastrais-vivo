# Plano de Refinamentos — Parser Vivo v2

## Contexto
O parser `parse_vivo.py` já processa com sucesso 36 arquivos `.txt`, gerando Excel + SQLite. Agora são necessários 3 refinamentos antes de considerar o pipeline pronto para produção:

1. **Deduplicação** inteligente (tudo igual exceto `arquivo_origem`)
2. **Planilha refinada visualmente** — cores alternadas, destaque para ativos, ordenação por data, negrito POS, mobile-friendly
3. **Geocodificação** — obter coordenadas e links de rota a partir do endereço completo

---

## 1. Deduplicação

### Regra de Negócio
Dois registros são considerados duplicados se **todas as colunas forem idênticas exceto `arquivo_origem`**.

### Implementação
```python
df_dedup = df.drop_duplicates(
    subset=[c for c in df.columns if c != "arquivo_origem"],
    keep="first"
)
```

### Onde inserir
- **Opção A (recomendada):** Deduplicar no DataFrame **antes** de exportar Excel/SQLite. Manter uma tabela `linhas_vivo_raw` (com duplicados) e `linhas_vivo` (deduplicada) no SQLite para auditoria.
- **Opção B:** Deduplicar apenas no Excel (para visualização) mas manter tudo no SQLite.

**Decisão:** Opção A. O usuário quer dados limpos; auditoria de duplicados pode ser feita via `COUNT(*) - COUNT(DISTINCT ...)` no raw.

---

## 2. Planilha Refinada (Visual/Mobile)

### Biblioteca
`openpyxl` (já é dependência via `pandas.ExcelWriter`). Não há skill específica de Excel/xlsx instalada no ambiente. A formatação será feita nativamente com `openpyxl.styles`.

### Especificações Visuais

| Requisito | Técnica openpyxl |
|-----------|------------------|
| **Cores alternadas (zebra)** | Loop nas linhas aplicando `PatternFill` cinza claro (`D9E1F2`) / branco alternados |
| **Destaque ATIVO** | Formatação condicional (ou pós-processamento): fundo verde claro (`C6EFCE`) + fonte verde escuro (`006100`) |
| **Prioridade data recente** | Ordenar DataFrame por `data_habilitacao_dt DESC` antes de escrever |
| **Negrito POS** | `Font(bold=True)` nas células da coluna `modalidade` quando valor == "POS" |
| **Mobile-friendly** | `ws.column_dimensions[col].auto_size = False` + larguras fixas razoáveis; `Alignment(wrap_text=True)`; congelar primeira linha (`freeze_panes`) |

### Detalhe de Implementação
```python
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active
ws.title = "Dados Cadastrais"

# Header
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF")

# Zebra
fill_white = PatternFill(fill_type=None)
fill_gray  = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

# Escrever dados já ordenados
for r_idx, row in enumerate(dataframe_to_rows(df_dedup, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=value)
        if r_idx == 1:
            cell.fill = header_fill
            cell.font = header_font
        else:
            situacao = df_dedup.iloc[r_idx-2]["situacao"]
            modalidade = df_dedup.iloc[r_idx-2]["modalidade"]
            if situacao == "ATIVO":
                cell.fill = fill_green
            else:
                cell.fill = fill_gray if (r_idx % 2 == 0) else fill_white
            if modalidade == "POS":
                cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
```

### Observação Mobile
O formato `.xlsx` é nativamente visualizável no Excel Mobile, Google Sheets e WPS. A chave para "bonito no celular" é:
- Larguras de coluna fixas e proporcionais (não `auto_size` que quebra no mobile)
- Texto com `wrap_text`
- Primeira linha congelada
- Cores de alto contraste (evitar azul escuro em fundo escuro)

---

## 3. Geocodificação (Coordenadas + Rotas)

### Serviço Escolhido: Nominatim (OpenStreetMap)
- **Gratuito**, sem API key obrigatória (rate limit: 1 req/s)
- Boa cobertura para endereços brasileiros
- Endpoint: `https://nominatim.openstreetmap.org/search?q={endereco}&format=json&limit=1`
- User-Agent obrigatório (política de uso)

### Alternativas consideradas
| Serviço | Custo | Precisão BR | Notas |
|---------|-------|-------------|-------|
| Nominatim | Grátis | Boa | Rate limit 1/s; requer User-Agent |
| Google Geocoding | Pago | Excelente | API key necessária; $5/1000 req |
| OpenCage | Freemium | Boa | API key; 2500 req/dia grátis |
| ViaCEP | Grátis | Só CEP | Não geocodifica endereço completo |

**Decisão:** Nominatim para MVP. Se precisar de maior volume/precisão, migrar para OpenCage ou Google.

### Arquitetura
1. **Concatenar endereço completo**: `f"{endereco}, {bairro}, {municipio}, {estado}, {cep}, Brazil"`
2. **Cache SQLite**: tabela `geocoding_cache` (`endereco_hash PRIMARY KEY`, `lat`, `lon`, `url_maps`, `data_consulta`)
3. **Batch com rate limit**: `time.sleep(1)` entre requisições; usar `requests.Session()`
4. **Link de rota**: Gerar URL do Google Maps Directions: `https://www.google.com/maps/dir/?api=1&destination={lat},{lon}`

### Schema Adicional
```sql
CREATE TABLE geocoding_cache (
    endereco_hash TEXT PRIMARY KEY,
    endereco_completo TEXT,
    latitude REAL,
    longitude REAL,
    google_maps_url TEXT,
    data_consulta TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

---

## 4. Skills Envolvidas

| Fase | Skill | Uso |
|------|-------|-----|
| Planejamento | `dev-lifecycle` | Fases 1-3 (requirements, design, planning) já foram internalizadas neste doc |
| Implementação | `python-patterns` | Type hints, dataclasses, EAFP, context managers para geocoding HTTP |
| Testes | `python-testing` | Fixtures para mock de Nominatim; testar deduplicação; testar estilos Excel |
| Verificação | `verify` | Re-executar parser + contar registros deduplicados + verificar formatação Excel |
| Documentação | `document-code` | Após implementação, documentar entry points e schema SQLite |

---

## 5. Checklist de Execução

- [ ] 1. Implementar `df.drop_duplicates()` no `parse_vivo.py` (exceto `arquivo_origem`)
- [ ] 2. Criar função `estilizar_excel()` usando openpyxl (cores, negrito, ordenação)
- [ ] 3. Ordenar DataFrame por `situacao` (ATIVO primeiro) + `data_habilitacao_dt DESC`
- [ ] 4. Implementar módulo `geocoding.py` com Nominatim + cache SQLite
- [ ] 5. Adicionar colunas `latitude`, `longitude`, `google_maps_url` ao DataFrame
- [ ] 6. Criar testes: deduplicação, estilo de célula ATIVO, mock de Nominatim
- [ ] 7. Rodar `verify`: parser → dedup → estilização → geocoding → assert counts
- [ ] 8. Gerar planilha final e validar visual no Excel/Google Sheets mobile

---

## 6. Anti-Patterns a Evitar

- **NÃO** fazer geocoding síncrono dentro do loop de parsing principal (separar em pipeline)
- **NÃO** usar `time.sleep(1)` se houver cache hit (só dormir após miss)
- **NÃO** salvar chaves de API no código (Nominatim não precisa, mas preparar para futuro)
- **NÃO** usar `openpyxl` diretamente no DataFrame se o pandas já estiver escrevendo; reabrir o workbook para estilizar

---

## Próximo Passo

Este plano está pronto para review. Após aprovação, a implementação segue a ordem do checklist acima.
