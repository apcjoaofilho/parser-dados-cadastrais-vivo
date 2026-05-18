# Parser de Dados Cadastrais Vivo

Ferramenta para extrair, consolidar, deduplicar, geocodificar e formatar dados cadastrais de linhas telefonicas Vivo a partir de relatorios em texto (.txt).

> **Aviso de Privacidade:** Os arquivos `.txt` contêm CPFs, nomes e endereços reais. **Nunca commite esses arquivos no GitHub.** Eles ficam automaticamente ignorados pelo `.gitignore` na pasta `data/`.
>
> **Seguranca dos arquivos de saida:** O Excel (`dados_cadastrais_vivo.xlsx`) e o banco SQLite (`dados_cadastrais.db`) gerados pelo programa tambem contêm dados pessoais. Proteja esses arquivos:
> - Nao os envie por e-mail ou mensageiros sem criptografia
> - Nao os deixe em pastas compartilhadas na rede
> - Apague-os quando nao forem mais necessarios
> - O cache de geocodificacao (`geocoding_cache` dentro do `.db`) tambem armazena endereços completos

---

## Para Usuarios Leigos (Nao Programadores)

### Instalacao (feita uma unica vez)

1. **Baixe este projeto:**
   - Clique no botao verde **<> Code** acima
   - Escolha **Download ZIP**
   - Extraia o ZIP para uma pasta facil de encontrar (ex: Area de Trabalho)

2. **Instale o Python (se ainda nao tiver):**
   - Va em [python.org/downloads](https://www.python.org/downloads/)
   - Baixe a ultima versao do **Python 3**
   - **IMPORTANTE:** Na tela de instalacao, MARQUE a opcao **"Add Python to PATH"**
   - Clique em **Install Now**

### Como Usar

1. Coloque todos os arquivos `.txt` da Vivo em uma pasta qualquer no seu computador.
2. Navegue ate a pasta do projeto, depois entre em `tools/vivo-parser/`.
3. De **duplo-clique** no arquivo:  
   **`Processar Dados Vivo.bat`**
4. O programa vai pedir:
   - **Pasta de entrada:** selecione onde estao os `.txt`
   - **Pasta de saida:** selecione onde quer salvar o resultado
5. Aguarde o processamento.
   - Na **primeira vez**, pode levar alguns minutos para instalar as bibliotecas necessarias.
   - Se a geocodificacao estiver ativada (padrao), o processamento demora mais porque consulta a internet para cada endereco.
6. Ao terminar, a pasta de saida sera aberta automaticamente.

### O que voce recebe

- **`dados_cadastrais_vivo.xlsx`** — Planilha Excel formatada, pronta para abrir no Excel, Google Sheets ou celular. Inclui:
  - Linhas ordenadas (ATIVO primeiro, depois por data)
  - Cores alternadas (zebra) para facilitar leitura
  - Destaque verde para registros **ATIVO**
  - Negrito na coluna modalidade quando for **POS**
  - Primeira linha congelada
  - Links do **Google Maps** para abrir rotas (quando a geocodificacao encontra o endereco)
- **`dados_cadastrais.db`** — Banco de dados SQLite com todas as informacoes (para quem quiser consultar com SQL).

### Duvidas? Veja a secao de Troubleshooting abaixo.

---

## Troubleshooting

### "Python nao foi encontrado"
Reinstale o Python de [python.org](https://www.python.org/downloads/) e certifique-se de marcar a opcao **"Add Python to PATH"** na tela de instalacao.

### O .bat fecha sozinho sem mostrar erro
Isso acontece se o Python nao estiver no PATH. Veja a mensagem acima.

### Antivirus bloqueia o .bat
Clique com o botao direito no arquivo e escolha **"Executar como administrador"**.

### Demora muito na primeira vez
Na primeira execucao, o programa instala automaticamente as bibliotecas `pandas`, `openpyxl` e `requests`. Isso pode levar 2-5 minutos dependendo da internet. Nas proximas vezes sera muito mais rapido.

### A geocodificacao esta muito lenta
A geocodificacao consulta o servico gratuito Nominatim (OpenStreetMap) com um limite de 1 requisicao por segundo. Para ~90 enderecos, leva cerca de 1 minuto e meio. Se quiser desativar, edite o arquivo `.bat` e adicione `--no-geocode` no final da linha que executa o parser.

### Erro "modulo nao encontrado" ou similar
Feche o terminal (pressione qualquer tecla se estiver pausado) e execute o `.bat` novamente. As vezes a instalacao das bibliotecas precisa de uma segunda tentativa.

---

## Para Desenvolvedores

### Dependencias

```bash
pip install -r tools/vivo-parser/requirements.txt
```

### Uso via Linha de Comando

```bash
python tools/vivo-parser/parse_vivo.py --input "caminho/para/txts" --output "caminho/para/saida"
```

Opcoes:
- `--no-geocode` — Desativa a geocodificacao (mais rapido, mas sem coordenadas/links do Maps)

### Estrutura do Projeto

```
.
├── .gitignore
├── README.md
├── data/                          # Dados de entrada (ignorado no git)
│   └── dados_cadastrais_vivo/
├── docs/                          # Documentacao
│   ├── plans/
│   ├── architecture/
│   └── specs/
├── output/                        # Pasta de saida padrao (ignorada no git, mantida via .gitkeep)
├── tools/
│   ├── vivo-parser/               # Parser Vivo
│   │   ├── Processar Dados Vivo.bat
│   │   ├── parse_vivo.py
│   │   ├── geocoding.py
│   │   ├── requirements.txt
│   │   └── tests/
│   │       └── test_parse_vivo.py
│   └── ficco_importer/            # Outro subprojeto
└── FICCO/                         # Vault Obsidian
```

### Executar Testes

```bash
python -m pytest tools/vivo-parser/tests/test_parse_vivo.py -v
```

### Funcionalidades

- **Parsing robusto:** Extrai campos de relatorios `.txt` com formatacao irregular, incluindo enderecos multilinha.
- **Deduplicacao inteligente:** Remove registros identicos ignorando a coluna `arquivo_origem`. Mantem tabela raw + deduplicada no SQLite para auditoria.
- **Planilha estilizada:** Cores zebra, destaque verde para ATIVO, negrito para POS, ordenacao por situacao e data, primeira linha congelada, larguras otimizadas para mobile.
- **Geocodificacao:** Usa Nominatim (OpenStreetMap) para obter latitude, longitude e gerar links do Google Maps. Inclui cache SQLite, busca em cascata para enderecos mal formatados e rate limit de 1 req/s.
- **Cache de geocodificacao:** Requisicoes sao cacheadas no mesmo arquivo `.db`, acelerando reexecucoes.

### Licenca

MIT
