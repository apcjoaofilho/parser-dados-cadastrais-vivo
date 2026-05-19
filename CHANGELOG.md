# Changelog

Todas as mudanças notáveis deste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/),
e este projeto adere ao [Semantic Versioning](https://semver.org/lang/pt-BR/).

## [2.0.0] - 2026-05-18

### Adicionado
- Parser completo de dados cadastrais Vivo (`tools/vivo-parser/parse_vivo.py`)
- Deduplicação automática de registros em múltiplos arquivos
- Geocodificação de endereços via Nominatim com cache SQLite
- Estilização avançada do Excel (zebra striping, ATIVO em verde, POS em negrito)
- Script batch para execução com duplo-clique (`Processar Dados Vivo.bat`)
- Módulos compartilhados em `tools/shared/` (pipeline.py, geocoding.py)
- Testes unitários (10 testes) e de integração (15 testes) — total 25/25 passando
- Validação de segurança contra path traversal e paths relativos
- Documentação completa em `docs/ai/implementation/knowledge-parse-vivo.md`
- Security review documentado em `docs/security/security-review-2026-05-18.md`

### Segurança
- Proteção contra path traversal (`..`)
- Paths relativos rejeitados
- Cache de geocoding documentado como contendo PII (endereços em texto plano)
- Pasta `data/` protegida por `.gitignore` (nunca commitada)

## [1.0.0] - 2026-05-17

### Adicionado
- Estrutura inicial do projeto
- Parser básico de dados cadastrais
- Geração de Excel simples

[2.0.0]: https://github.com/apcjoaofilho/parser-dados-cadastrais-vivo/releases/tag/v2.0.0
[1.0.0]: https://github.com/apcjoaofilho/parser-dados-cadastrais-vivo/releases/tag/v1.0.0
