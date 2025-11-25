# PDF→Excel Ingestor

Ferramenta genérica para **extrair dados estruturados de PDFs** e gerar planilhas **Excel** seguindo um **modelo XLSX pré-existente**.

Ideal para automações onde PDFs seguem um padrão visual/tabelado e precisam ser convertidos em linhas de uma planilha (por exemplo: fichas de cadastro, documentos corporativos, registros, formulários digitalizados etc.).

---

## ✨ Funcionalidades

- Extração de texto via **pdfplumber**  
- Suporte a PDF digital e PDF escaneado (OCR com Tesseract)  
- Mapeamento configurável por **YAML**  
- Geração de planilha Excel utilizando um **template base**  
- Fallback OCR para páginas com baixa qualidade  
- Logs detalhados por arquivo  
- Execução em lote (vários PDFs de uma vez)  
- Preserva o layout do XLSX original  
- Flags de depuração e auditoria

---

## 🛠️ Tecnologias utilizadas

- Python 3.10+
- pdfplumber
- PyYAML
- openpyxl
- pdf2image (Poppler)
- pytesseract (Tesseract OCR)
- pillow
- opencv-python

---

## 📁 Estrutura do projeto

/
├── run.py # Runner principal (CLI)
├── pdf_excel_ingestor.py # Motor principal de extração e escrita
├── mapping.yaml # Configuração de mapeamento dos campos
├── MODELO_PLANILHA_INCLUSAO.xlsx # Template de saída (não versionado)
├── entrada/ # PDFs de entrada
├── saida/ # XLSX gerados
├── requirements.txt
└── README.md

yaml
Copiar código

---

## 🔧 Instalação

### 1. Criar ambiente virtual

```bash
python -m venv .venv
.\.venv\Scripts\activate
2. Instalar dependências
bash
Copiar código
pip install -r requirements.txt
3. Instalar dependências externas
Windows
Instalar Tesseract OCR:
https://github.com/tesseract-ocr/tesseract

Instalar Poppler (necessário para pdf2image):
https://github.com/oschwartz10612/poppler-windows/releases/

Adicionar ambos ao PATH.

🚀 Uso básico
Coloque seus PDFs dentro da pasta entrada/ e rode:

bash
Copiar código
py run.py
Se houver template XLSX na raiz com o nome:

Copiar código
MODELO_PLANILHA_INCLUSAO.xlsx
ele será detectado automaticamente.

🚀 Exemplo com argumentos
Rodar somente PDFs dentro de uma pasta específica:

bash
Copiar código
py run.py -i "entrada_lote/*.pdf"
Usar um template específico:

bash
Copiar código
py run.py -t "MEU_TEMPLATE.xlsx"
Alterar o nome final da planilha:

bash
Copiar código
py run.py --xlsx-name "resultado_final.xlsx"
Ativar OCR em todas as páginas:

bash
Copiar código
py run.py --force-ocr
Debug:

bash
Copiar código
py run.py --debug-argv
🧩 Como funciona o mapeamento (mapping.yaml)
O arquivo mapping.yaml define como os campos extraídos do PDF serão transferidos para colunas específicas do Excel.

Exemplo simplificado:

yaml
Copiar código
beneficiario_nome:
  regex: "Nome completo: (.*)"
  column: "B2"

cpf:
  regex: "CPF: ([0-9\\.\\-]+)"
  column: "C2"
Você pode adicionar, remover ou adaptar campos conforme a necessidade do layout.

🏗️ Template XLSX
O template deve conter:

Estrutura final desejada

Cabeçalhos

Fórmulas

Formatação

Colunas/espaços predefinidos

O script não modifica a formatação — ele preenche exatamente nas células definidas.

📜 Logs
Os logs são exibidos no console, ex.:

mathematica
Copiar código
INFO  | Processando 58 PDF(s)...
INFO  | PDF: ficha_01.pdf
INFO  | PDF: ficha_02.pdf
...
❗ Possíveis erros
Erro: "Informe --template"
O template XLSX não foi encontrado.
Você deve:

colocar MODELO_PLANILHA_INCLUSAO.xlsx na raiz
ou

informar manualmente:

bash
Copiar código
py run.py -t "meu_template.xlsx"
Erro: Nenhum PDF encontrado
A pasta entrada/ está vazia.

📄 Licença
MIT – uso livre para projetos pessoais ou comerciais.

🤝 Contribuindo
Pull requests são bem-vindos.
Para grandes alterações, abra primeiro uma issue para discutir o que deseja alterar.

👤 Autor
Desenvolvido por Luis Felipe Brandt Barbosa
GitHub: https://github.com/lfbrandt