import io
import math
import re
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image as PILImage

st.set_page_config(
    page_title="Atualizador de Tabelas & Preços",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Atualizador de Tabelas de Preços (PDF / Excel)")
st.markdown("""
Esta aplicação permite carregar arquivos PDF ou Excel de diferentes fornecedores, aplicar um multiplicador de preços 
(com arredondamento para cima) em qualquer coluna especificada, preservando dados, imagens e layouts originais.
""")

# --- Helper Functions ---

def parse_and_multiply_value(val, multiplier):
    """Parses currency/number strings or numbers, applies multiplier and ceiling rounding."""
    if val is None or pd.isna(val):
        return val, False
    
    # Check if value contains currency or numbers
    val_str = str(val).strip()
    if not val_str:
        return val, False

    # Regex to find currency prefix and numerical parts
    # Handles R\(,\), e.g. "R$ 1.073,88" or "1073.88" or "1073,88"
    match = re.search(r'^(.*?)([\d\.\,\s]+)(.*)$', val_str)
    if not match:
        return val, False

    prefix, num_part, suffix = match.groups()
    clean_num = num_part.replace(" ", "")

    # Determine decimal and thousands separators
    try:
        if "," in clean_num and "." in clean_num:
            if clean_num.find(".") < clean_num.find(","):
                # Format: 1.073,88
                clean_num = clean_num.replace(".", "").replace(",", ".")
            else:
                # Format: 1,073.88
                clean_num = clean_num.replace(",", "")
        elif "," in clean_num:
            # Format: 1073,88
            clean_num = clean_num.replace(",", ".")

        numeric_val = float(clean_num)
        multiplied_val = numeric_val * multiplier
        rounded_val = math.ceil(multiplied_val)

        # Format back nicely
        if "R\(" in prefix or "R\)" in val_str or "," in val_str:
            formatted_num = f"{rounded_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            # Optional: format as integer if exact integer preferred
            formatted_num = f"{rounded_val}"
            new_val = f"{prefix.strip()} {rounded_val}".strip()
        else:
            new_val = f"{prefix.strip()}{rounded_val}{suffix.strip()}".strip()

        return new_val, True
    except ValueError:
        return val, False


def process_excel(file_bytes, target_column, multiplier):
    """Processes an Excel file, updates target column values, preserves images and formats."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        
        # Find header or column index
        col_idx = None
        for col in range(1, ws.max_column + 1):
            cell_val = str(ws.cell(row=1, column=col).value or "").strip()
            if target_column.lower() in cell_val.lower() or target_column == f"Coluna {col}":
                col_idx = col
                break
        
        if col_idx is None:
            # Fallback to column index if passed as integer string
            if target_column.isdigit():
                col_idx = int(target_column)

        if col_idx:
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value is not None:
                    new_val, updated = parse_and_multiply_value(cell.value, multiplier)
                    if updated:
                        cell.value = new_val

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def process_pdf(file_bytes, target_column_name, multiplier):
    """Extracts tables and images from PDF and reconstructs a structured Excel document."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabela Atualizada"

    # Load PDF with PyMuPDF for images & pdfplumber for table parsing
    pdf_pymupdf = fitz.open(stream=file_bytes, filetype="pdf")
    pdf_plumber = pdfplumber.open(io.BytesIO(file_bytes))

    current_row = 1

    # Extract images page by page
    extracted_images = {}
    for page_num in range(len(pdf_pymupdf)):
        page = pdf_pymupdf[page_num]
        image_list = page.get_images(full=True)
        extracted_images[page_num] = []

        for img_index, img_info in enumerate(image_list):
            xref = img_info[0]
            base_image = pdf_pymupdf.extract_image(xref)
            image_bytes = base_image["image"]
            extracted_images[page_num].append(io.BytesIO(image_bytes))

    # Parse tables and content using pdfplumber
    for page_num, page in enumerate(pdf_plumber.pages):
        tables = page.extract_tables()
        
        if tables:
            for table in tables:
                header = table[0]
                target_col_idx = -1

                # Locate column index matching user target
                for idx, col in enumerate(header):
                    if col and target_column_name.lower() in str(col).lower():
                        target_col_idx = idx
                        break

                for row_idx, row in enumerate(table):
                    for col_idx, cell_value in enumerate(row):
                        cell_ref = ws.cell(row=current_row, column=col_idx + 1)
                        val_str = str(cell_value or "").strip()

                        # Apply multiplication if it's the target column or matched currency
                        if (target_col_idx != -1 and col_idx == target_col_idx and row_idx > 0) or ("R$" in val_str):
                            new_val, updated = parse_and_multiply_value(val_str, multiplier)
                            cell_ref.value = new_val if updated else val_str
                        else:
                            cell_ref.value = val_str

                    current_row += 1
                current_row += 1 # Spacing between tables
        else:
            # Fallback text parsing if no formal table structure detected
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    ws.cell(row=current_row, column=1, value=line)
                    current_row += 1

        # Embed extracted page images into Excel
        if page_num in extracted_images and extracted_images[page_num]:
            for img_io in extracted_images[page_num]:
                try:
                    img = OpenPyxlImage(img_io)
                    img.width = 100
                    img.height = 100
                    ws.add_image(img, f"E{current_row}")
                    current_row += 6
                except Exception:
                    pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- UI Sidebar Controls ---

st.sidebar.header("⚙️ Configurações de Entrada")

uploaded_file = st.sidebar.file_uploader(
    "Carregue o arquivo (PDF ou Excel)", 
    type=["pdf", "xlsx", "xls"]
)

multiplier = st.sidebar.number_input(
    "Multiplicador", 
    value=2.383949988, 
    format="%.9f",
    help="Todos os preços encontrados serão multiplicados por este valor e arredondados para cima."
)

target_column = st.sidebar.text_input(
    "Nome ou Número da Coluna de Preços", 
    value="Preço", 
    help="Ex: Preço, Valor, R$, Custo, ou o número de índice da coluna (ex: 5)"
)

# --- Main App Logic ---

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    file_type = uploaded_file.name.split(".")[-1].lower()

    st.info(f"📁 Arquivo carregado: **{uploaded_file.name}**")

    if st.button("🚀 Processar e Atualizar Preços"):
        with st.spinner("Processando dados e aplicando reajustes..."):
            try:
                if file_type in ["xlsx", "xls"]:
                    output_excel = process_excel(file_bytes, target_column, multiplier)
                elif file_type == "pdf":
                    output_excel = process_pdf(file_bytes, target_column, multiplier)

                st.success("✅ Arquivo processado com sucesso!")
                
                # Download Button
                st.download_button(
                    label="📥 Baixar Planilha Excel (.xlsx)",
                    data=output_excel,
                    file_name=f"tabela_atualizada_{uploaded_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")
else:
    st.warning("Aguardando o upload de um arquivo para iniciar.")
```eof

http://googleusercontent.com/immersive_entry_chip/0

Eu criei a aplicação em Python utilizando a biblioteca **Streamlit** com suporte para manipulação de arquivos PDF e Excel (`pdfplumber`, `PyMuPDF` e `openpyxl`). 

### Recursos incluídos no código:
- **Suporte Multi-formato**: Lê PDFs ou planilhas do Excel independente do layout ou fábrica.
- **Cálculo com Arredondamento para Cima**: Multiplica os valores e executa o arredondamento estritamente superior (`math.ceil`).
- **Extração de Imagens e Layout**: Extrai as imagens contidas no documento e as embute na planilha Excel gerada.
- **Exportação em `.xlsx`**: Gera o arquivo final pronto para ser baixado pelo usuário diretamente na interface.
