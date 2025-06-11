# pdf_reader

Este projeto tem como objetivo automatizar a extração e organização de dados a partir de ficheiros PDF, convertendo-os em formatos estruturados como JSON e Excel, e facilitando a análise de itens presentes nesses documentos.

## Organização do Projeto

pdf_reader/
├── src/ # Código principal da aplicação
├── tests/ # Scripts de teste e experimentação
├── excels_from_pdf/ # Resultados: ficheiros Excel extraídos dos PDFs
├── excels_visual/ # Ficheiros Excel gerados a partir dos JSONs
├── jsons/ # Ficheiros JSON gerados a partir dos PDFs
├── pdfs/ # PDFs de entrada
├── notepads/ # Ficheiros txt gerados a partir dos PDFs
├── package.json
├── package-lock.json
├── .gitignore
├── README.md

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

    - pdf_to_json.py: Para converter PDF em JSON

    - json_to_excel.py: Para converter JSON em Excel

5. Os resultados serão gravados nas respetivas pastas de saída.


## Dependências

**Para Node.js:**

    - npm install

**Para Python, recomenda-se o uso de um ambiente virtual:**

    - python -m venv plumber_env
    - source plumber_env/bin/activate  # ou plumber_env\Scripts\activate no Windows
    - pip install -r requirements.txt  # (se tiveres este ficheiro)
    