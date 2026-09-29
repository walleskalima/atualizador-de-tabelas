import io
import math
import re
import tempfile
from decimal import Decimal, ROUND_CEILING
from pathlib import Path

import pandas as pd
import streamlit as st
import pymupdf
from PIL import Image as PILImage
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill


st.set_page_config(
    page_title="Atualizador de Tabelas",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Atualizador de preços")
st.write(
    "Envie uma tabela Excel ou PDF, informe o multiplicador, "
    "escolha a coluna de preços e baixe o Excel atualizado."
)


def arredondar_para_cima(valor):
    return int(
        Decimal(str(valor)).quantize(
            Decimal("1"),
            rounding=ROUND_CEILING,
        )
    )


def converter_texto_para_numero(valor):
    if valor is None:
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.replace("R$", "").replace("r$", "").strip()
    texto = texto.replace(" ", "")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    texto = re.sub(r"[^0-9.\-]", "", texto)

    if not texto:
        return None

    try:
        return float(texto)
    except ValueError:
        return None


def formatar_moeda(valor):
    return f"R$ {valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parece_preco(valor):
    if isinstance(valor, (int, float)):
        return True

    if not isinstance(valor, str):
        return False

    texto = valor.strip().lower()

    if "r$" in texto:
        return True

    padrao = r"^\s*\d{1,3}(\.\d{3})*(,\d{1,2})?\s*$"
    return bool(re.match(padrao, texto))


def atualizar_valor(valor_original, multiplicador):
    numero = converter_texto_para_numero(valor_original)

    if numero is None:
        return None

    novo_valor = arredondar_para_cima(numero * multiplicador)

    if isinstance(valor_original, str):
        if "R$" in valor_original.upper():
            return formatar_moeda(novo_valor)

        if "," in valor_original:
            return f"{novo_valor:,.0f}".replace(",", ".")

        return str(novo_valor)

    return novo_valor


def ler_cabecalhos(planilha, linha_cabecalho):
    cabecalhos = {}

    for coluna in range(1, planilha.max_column + 1):
        valor = planilha.cell(linha_cabecalho, coluna).value

        if valor is not None and str(valor).strip():
            cabecalhos[str(valor).strip()] = coluna

    return cabecalhos


def atualizar_excel(
    arquivo_bytes,
    nome_aba,
    linha_cabecalho,
    numero_coluna,
    multiplicador,
):
    workbook = load_workbook(
        io.BytesIO(arquivo_bytes),
        data_only=False,
        keep_links=True,
        rich_text=True,
    )

    planilha = workbook[nome_aba]
    alteracoes = []

    for linha in range(linha_cabecalho + 1, planilha.max_row + 1):
        celula = planilha.cell(linha, numero_coluna)
        valor_original = celula.value

        if valor_original is None:
            continue

        if isinstance(valor_original, str) and valor_original.startswith("="):
            continue

        if not parece_preco(valor_original):
            continue

        novo_valor = atualizar_valor(valor_original, multiplicador)

        if novo_valor is None or novo_valor == valor_original:
            continue

        celula.value = novo_valor

        if isinstance(novo_valor, (int, float)):
            formato = celula.number_format or ""

            if "R$" in formato or "[$R$" in formato:
                celula.number_format = 'R$ #,##0'

        alteracoes.append(
            {
                "Linha": linha,
                "Célula": celula.coordinate,
                "Original": valor_original,
                "Atualizado": novo_valor,
            }
        )

    arquivo_saida = io.BytesIO()
    workbook.save(arquivo_saida)
    arquivo_saida.seek(0)

    return arquivo_saida.getvalue(), alteracoes


def atualizar_precos_no_texto(texto, multiplicador):
    padrao = r"R\$\s*([\d\.]+,\d{2})"

    def substituir(match):
        numero = converter_texto_para_numero(match.group(1))

        if numero is None:
            return match.group(0)

        novo_valor = arredondar_para_cima(numero * multiplicador)
        return formatar_moeda(novo_valor)

    return re.sub(padrao, substituir, texto)


def inserir_pagina_como_imagem(planilha, pagina):
    pix = pagina.get_pixmap(
        matrix=pymupdf.Matrix(1.5, 1.5),
        alpha=False,
    )

    imagem = PILImage.open(io.BytesIO(pix.tobytes("png")))

    arquivo_temporario = tempfile.NamedTemporaryFile(
        suffix=".png",
        delete=False,
    )

    imagem.save(arquivo_temporario.name, format="PNG")
    arquivo_temporario.close()

    imagem_excel = ExcelImage(arquivo_temporario.name)

    largura_maxima = 1100

    if imagem_excel.width > largura_maxima:
        proporcao = largura_maxima / imagem_excel.width
        imagem_excel.width = int(imagem_excel.width * proporcao)
        imagem_excel.height = int(imagem_excel.height * proporcao)

    planilha.add_image(imagem_excel, "A1")
    planilha.column_dimensions["A"].width = 20


def converter_pdf_para_excel(arquivo_bytes, multiplicador):
    documento = pymupdf.open(
        stream=arquivo_bytes,
        filetype="pdf",
    )

    workbook = Workbook()
    planilha_texto = workbook.active
    planilha_texto.title = "Tabela editável"

    titulos = ["Página", "Texto original", "Texto com valores atualizados"]

    for coluna, titulo in enumerate(titulos, start=1):
        celula = planilha_texto.cell(1, coluna, titulo)
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")
        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    planilha_texto.column_dimensions["A"].width = 12
    planilha_texto.column_dimensions["B"].width = 70
    planilha_texto.column_dimensions["C"].width = 70

    linha_excel = 2
    total_alteracoes = 0

    for numero_pagina, pagina in enumerate(documento, start=1):
        texto_pagina = pagina.get_text("text")
        linhas = [
            linha.strip()
            for linha in texto_pagina.splitlines()
            if linha.strip()
        ]

        for texto in linhas:
            texto_atualizado = atualizar_precos_no_texto(
                texto,
                multiplicador,
            )

            if texto_atualizado != texto:
                total_alteracoes += 1

            planilha_texto.cell(linha_excel, 1, numero_pagina)
            planilha_texto.cell(linha_excel, 2, texto)
            planilha_texto.cell(linha_excel, 3, texto_atualizado)

            for coluna in range(1, 4):
                planilha_texto.cell(
                    linha_excel,
                    coluna,
                ).alignment = Alignment(
                    wrap_text=True,
                    vertical="top",
                )

            linha_excel += 1

        planilha_visual = workbook.create_sheet(
            f"Visual - Página {numero_pagina}"
        )
        inserir_pagina_como_imagem(planilha_visual, pagina)

    planilha_info = workbook.create_sheet("Informações")
    planilha_info["A1"] = "Conversão de PDF para Excel"
    planilha_info["A1"].font = Font(bold=True, size=14)

    planilha_info["A3"] = "Multiplicador aplicado"
    planilha_info["B3"] = multiplicador

    planilha_info["A4"] = "Arredondamento"
    planilha_info["B4"] = "Sempre para cima"

    planilha_info["A6"] = "Observação"
    planilha_info["B6"] = (
        "A aba 'Tabela editável' traz o conteúdo extraído e atualizado. "
        "As abas 'Visual - Página N' mantêm cada página do PDF como imagem, "
        "preservando o visual, fotos, logos e disposição geral."
    )

    planilha_info.column_dimensions["A"].width = 28
    planilha_info.column_dimensions["B"].width = 100
    planilha_info["B6"].alignment = Alignment(wrap_text=True)

    arquivo_saida = io.BytesIO()
    workbook.save(arquivo_saida)
    arquivo_saida.seek(0)

    return arquivo_saida.getvalue(), total_alteracoes


arquivo = st.file_uploader(
    "Envie uma tabela Excel ou PDF",
    type=["xlsx", "pdf"],
)

multiplicador = st.number_input(
    "Valor para multiplicar os preços",
    min_value=0.000001,
    value=2.383949988,
    step=0.000001,
    format="%.9f",
)

if arquivo is None:
    st.info("Envie um arquivo para começar.")
    st.stop()

extensao = Path(arquivo.name).suffix.lower()
conteudo_arquivo = arquivo.getvalue()

if extensao == ".xlsx":
    workbook_preview = load_workbook(
        io.BytesIO(conteudo_arquivo),
        data_only=False,
        keep_links=True,
        rich_text=True,
    )

    aba = st.selectbox(
        "Escolha a aba da tabela",
        workbook_preview.sheetnames,
    )

    planilha_preview = workbook_preview[aba]

    linha_cabecalho = st.number_input(
        "Em qual linha estão os títulos das colunas?",
        min_value=1,
        max_value=max(1, planilha_preview.max_row),
        value=1,
        step=1,
    )

    cabecalhos = ler_cabecalhos(
        planilha_preview,
        linha_cabecalho,
    )

    if not cabecalhos:
        st.error(
            "Não foi possível identificar os títulos nessa linha. "
            "Tente escolher outra linha."
        )
        st.stop()

    nome_coluna = st.selectbox(
        "Escolha a coluna de preço/valor",
        list(cabecalhos.keys()),
    )

    numero_coluna = cabecalhos[nome_coluna]

    if st.button("Gerar Excel atualizado", type="primary"):
        with st.spinner("Atualizando valores..."):
            resultado, alteracoes = atualizar_excel(
                conteudo_arquivo,
                aba,
                linha_cabecalho,
                numero_coluna,
                multiplicador,
            )

        st.success(
            f"Pronto. Foram atualizados {len(alteracoes)} valor(es)."
        )

        if alteracoes:
            st.dataframe(
                pd.DataFrame(alteracoes),
                use_container_width=True,
                hide_index=True,
            )

        nome_saida = f"{Path(arquivo.name).stem}_atualizado.xlsx"

        st.download_button(
            label="⬇️ Baixar Excel atualizado",
            data=resultado,
            file_name=nome_saida,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

elif extensao == ".pdf":
    st.warning(
        "PDFs podem ter tabelas, fotos, logos e textos em posições diferentes. "
        "O resultado terá uma aba editável e abas visuais com as páginas originais."
    )

    if st.button("Converter PDF e atualizar valores", type="primary"):
        with st.spinner("Convertendo o PDF..."):
            resultado, alteracoes = converter_pdf_para_excel(
                conteudo_arquivo,
                multiplicador,
            )

        st.success(
            f"Pronto. Foram identificadas {alteracoes} linha(s) com valores monetários."
        )

        nome_saida = f"{Path(arquivo.name).stem}_convertido.xlsx"

        st.download_button(
            label="⬇️ Baixar Excel convertido",
            data=resultado,
            file_name=nome_saida,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
