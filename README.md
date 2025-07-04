# pdf_reader

Este projeto tem como objetivo automatizar a extração e organização de dados a partir de ficheiros PDF, convertendo-os em formatos estruturados como JSON e Excel, e facilitando a análise de itens presentes nesses documentos.

## Organização do Projeto

pdf_reader/
├── src/                    # Código principal da aplicação
│
├── tests/                 # Scripts de teste e experimentação
│   └── src_v1/            # Primeira versão do projeto (legado ou protótipo inicial)
│                  
├── pdfs/                  # PDFs de entrada
│
├── jsons/                 # Resultado: Ficheiros JSON gerados a partir dos PDFs
├── excels_raw/            # Resultado: Excel gerado diretamente dos PDFs
├── excels_processed/      # Resultado: Excel gerado a partir dos JSONs processados
├── notepads/              # Resultado: Ficheiros .txt gerados dos PDFs (textosimples)
│
│                          # OBS: Estas pastas de "Resultado" não existem inicialmente,
│                          #      mas são criadas automaticamente quando o projeto é executado.
│
├── .gitignore             # Ficheiros/pastas a ignorar no controlo de versões (Git)
├── package.json           # Dependências e scripts do projeto Node.js
├── package-lock.json      # Versões exatas das dependências instaladas
└── README.md              # Documentação inicial do projeto


### Descrição das Pastas

- **`src/`**  
  Contém o **código principal do projeto**. O fluxo principal é:
  1. Conversão do PDF para JSON
  2. Conversão do JSON para Excel
  3. Leitura do Excel para identificar e extrair cada item, incluindo a quantidade e informações associadas a cada linha do ficheiro.

  Os scripts nesta pasta representam a versão consolidada e limpa do programa, prontos para utilização.

- **`tests/`**  
  Inclui **os vários scripts e ficheiros usados durante o desenvolvimento** para testar funcionalidades, experimentar abordagens e compreender o funcionamento da leitura e organização dos PDFs. São rascunhos, protótipos e exemplos intermédios que serviram de apoio à construção da solução principal mas ainda funcionais.

- **`result/`, `excels_from_pdf/`, `excels_visual/`, `jsons/`, `pdfs/`, `notepads/`**  
  Estas pastas são criadas automaticamente pelo programa à medida que vais correndo os scripts e gerando resultados (Excel, JSON, Notepad, etc.).
  Por isso, podem não aparecer logo após clonares o repositório — surgem quando necessário, durante a utilização normal.
  A única pasta que precisas de criar manualmente é a pdfs/, onde deves colocar os PDFs de entrada que queres processar.


## Como Usar

1. **Clona o repositório:**

   git clone <url-do-repositorio>

2. Cria uma pasta pdfs/.

3. Adiciona os teus PDFs na pasta pdfs/.

4. Corre os scripts principais que estão em src/ conforme a tua necessidade:

    - pdf_to_json.py: Responsável por converter os ficheiros PDF em ficheiros JSON.

    - json_reader.py: Responsável por ler os ficheiros JSON e extrair as informações relevantes.

    - auto_reader.py: Executa automaticamente os processos de conversão e leitura combinando os dois ficheiros anteriores.

5. Os resultados serão gravados nas respetivas pastas de saída.


## Dependências

**Para Node.js:**

    - npm install

**Para auto-file.py**

    - pip install pandas
    - pip install pymupdf
    - pip install XlsxWriter
    - pip install openpyxl


**Para Python, recomenda-se o uso de um ambiente virtual:**

    - python -m venv plumber_env
    - source plumber_env/bin/activate  # ou plumber_env\Scripts\activate no Windows
    - pip install -r requirements.txt  # (se tiveres este ficheiro)

**Última atualização:** 04 de julho de 2025