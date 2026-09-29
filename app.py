import io
import re
from decimal import Decimal, InvalidOperation, ROUND_CEILING
from pathlib import Path

import pandas as pd
import pymupdf
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(
    page_title="Atualizador de Tabelas",
    page_icon="💰",
    layout="wide",
)

st.title("💰 Atualizador de preços")
st.caption(
    "PDF entra como PDF e sai como PDF. "
    "Excel entra como Excel e sai como Excel."
)


# ==========================================================
# CÁLCULO E FORMATAÇÃO DE PREÇOS
# ==========================================================

PADRAO_PRECO = re.compile(
    r"R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}"
)


def arredondar_sempre_para_cima(valor):
    """
    Exemplos:
    10,01 -> 11
    10,99 -> 11
    20,00 -> 20
    """
    return int(
        Decimal(str(valor)).quantize(
            Decimal("1"),
            rounding=ROUND_CEILING,
        )
    )


def texto_preco_para_decimal(texto_preco):
    """
    Converte:
    R$ 926,82 -> Decimal('926.82')
    R$ 1.073,88 -> Decimal('1073.88')
    """
    texto = texto_preco.upper().replace("R$", "").strip()
    texto = texto.replace(".", "").replace(",", ".")

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def formatar_preco_brasileiro(valor_inteiro):
    """
    Converte:
    2210 -> R$ 2.210
    2561 -> R$ 2.561
    """
    numero = f"{valor_inteiro:,}"
    numero = numero.replace(",", ".")
    return f"R$ {numero}"


def calcular_novo_preco(texto_preco, multiplicador):
    """
    Lê um preço no formato brasileiro, multiplica e arredonda para cima.
    """
    preco_original = texto_preco_para_decimal(texto_preco)

    if preco_original is None:
        return None

    resultado = preco_original * Decimal(str(multiplicador))
    preco_final = arredondar_sempre_para_cima(resultado)

    return preco_final


# ==========================================================
# PDF -> PDF
# ==========================================================

def rgb_inteiro_para_tupla(cor):
    """
    O PyMuPDF pode entregar cor como inteiro RGB.
    Converte para uma tupla entre 0 e 1.
    """
    if not isinstance(cor, int):
        return (0, 0, 0)

    vermelho = ((cor >> 16) & 255) / 255
    verde = ((cor >> 8) & 255) / 255
    azul = (cor & 255) / 255

    return (vermelho, verde, azul)


def buscar_spans_com_preco(pagina):
    """
    Procura preços dentro dos spans reais de texto do PDF.

    Retorna:
    - o texto do preço;
    - a caixa de posição dele;
    - tamanho de fonte;
    - cor;
    - nome de fonte;
    """
    encontrados = []

    estrutura = pagina.get_text("dict")

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue

        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                texto_span = span.get("text", "")

                if not texto_span:
                    continue

                for correspondencia in PADRAO_PRECO.finditer(texto_span):
                    texto_preco = correspondencia.group(0)

                    # Na grande maioria de catálogos, o preço vem em um span próprio.
                    # Quando estiver dentro de um span maior, estima a posição do trecho.
                    x0, y0, x1, y1 = span["bbox"]

                    inicio = correspondencia.start()
                    fim = correspondencia.end()
                    tamanho_total = max(len(texto_span), 1)

                    if inicio == 0 and fim == len(texto_span):
                        caixa_preco = pymupdf.Rect(x0, y0, x1, y1)
                    else:
                        largura_span = x1 - x0
                        x_inicio = x0 + (largura_span * inicio / tamanho_total)
                        x_fim = x0 + (largura_span * fim / tamanho_total)

                        caixa_preco = pymupdf.Rect(
                            x_inicio,
                            y0,
                            x_fim,
                            y1,
                        )

                    encontrados.append(
                        {
                            "texto_original": texto_preco,
                            "rect": caixa_preco,
                            "fonte": span.get("font", "helv"),
                            "tamanho": span.get("size", 10),
                            "cor": rgb_inteiro_para_tupla(
                                span.get("color", 0)
                            ),
                        }
                    )

    return encontrados


def escolher_fonte_pdf(nome_fonte_original):
    """
    Nem toda fonte incorporada no PDF pode ser reutilizada diretamente.
    Usa fontes-base seguras do PDF.

    Para preservar aparência:
    - fonte com 'bold' usa Helvetica Bold;
    - fonte com 'times' usa Times Roman;
    - demais usam Helvetica.
    """
    nome = str(nome_fonte_original).lower()

    if "bold" in nome:
        return "hebo"

    if "times" in nome:
        return "tiro"

    if "courier" in nome:
        return "cour"

    return "helv"


def largura_texto(pagina, texto, fonte, tamanho):
    """
    Mede a largura real do novo texto no PDF.
    """
    try:
        return pymupdf.get_text_length(
            texto,
            fontname=fonte,
            fontsize=tamanho,
        )
    except Exception:
        return len(texto) * tamanho * 0.55


def atualizar_pdf_mantendo_layout(arquivo_pdf, multiplicador):
    """
    Atualiza preços diretamente no PDF.

    Estratégia:
    1. Localiza cada preço como texto.
    2. Registra posição, tamanho e cor.
    3. Faz uma redação branca somente no retângulo do preço.
    4. Aplica a redação para remover de fato o texto antigo.
    5. Insere o novo preço na mesma posição.
    """
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    for pagina in documento:
    pagina.draw_rect(
        pagina.rect,
        color=None,
        fill=(1, 1, 1),
        overlay=False,
    )

    conferencias = []

    for indice_pagina, pagina in enumerate(documento, start=1):
        precos_pagina = buscar_spans_com_preco(pagina)
        substituicoes = []

        for item in precos_pagina:
            preco_original_texto = item["texto_original"]
            preco_final = calcular_novo_preco(
                preco_original_texto,
                multiplicador,
            )

            if preco_final is None:
                continue

            texto_novo = formatar_preco_brasileiro(preco_final)

            caixa_original = item["rect"]
            margem_x = 1.2
            margem_y = 0.8

            # A área é expandida só um pouco para não deixar restos do preço antigo.
            caixa_redacao = pymupdf.Rect(
                caixa_original.x0 - margem_x,
                caixa_original.y0 - margem_y,
                caixa_original.x1 + margem_x,
                caixa_original.y1 + margem_y,
            )

            # A redação remove o texto do preço antigo.
            # O preenchimento branco funciona para páginas de fundo branco/claro.
            pagina.add_redact_annot(
                caixa_redacao,
                fill=(1, 1, 1),
            )

            substituicoes.append(
                {
                    "texto_original": preco_original_texto,
                    "texto_novo": texto_novo,
                    "caixa_original": caixa_original,
                    "caixa_redacao": caixa_redacao,
                    "fonte": escolher_fonte_pdf(item["fonte"]),
                    "tamanho": item["tamanho"],
                    "cor": item["cor"],
                }
            )

        if not substituicoes:
            continue

        # Remove definitivamente apenas o texto dentro das caixas de preço.
        # As imagens e vetores fora dessas caixas continuam no documento.
        pagina.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )

        for item in substituicoes:
            caixa_original = item["caixa_original"]
            texto_novo = item["texto_novo"]
            fonte = item["fonte"]
            tamanho = item["tamanho"]
            cor = item["cor"]

            largura_original = caixa_original.width
            largura_nova = largura_texto(
                pagina,
                texto_novo,
                fonte,
                tamanho,
            )

            # Mantém alinhamento:
            # se o novo texto couber, conserva a posição à esquerda;
            # se for maior, reduz até caber na área original.
            tamanho_ajustado = tamanho

            while (
                largura_nova > largura_original
                and tamanho_ajustado > 5
            ):
                tamanho_ajustado -= 0.2
                largura_nova = largura_texto(
                    pagina,
                    texto_novo,
                    fonte,
                    tamanho_ajustado,
                )

            # Caixa levemente mais larga para não cortar o texto,
            # mas mantém a mesma altura e linha do preço.
            caixa_texto = pymupdf.Rect(
                caixa_original.x0,
                caixa_original.y0 - 0.5,
                max(
                    caixa_original.x1,
                    caixa_original.x0 + largura_nova + 2,
                ),
                caixa_original.y1 + 1.5,
            )

            # insert_textbox mantém o texto dentro da área da posição original.
            resultado = pagina.insert_textbox(
                caixa_texto,
                texto_novo,
                fontname=fonte,
                fontsize=tamanho_ajustado,
                color=cor,
                align=pymupdf.TEXT_ALIGN_LEFT,
                overlay=True,
            )

            # Caso o textbox fique muito justo, usa escrita direta.
            if resultado < 0:
                pagina.insert_text(
                    pymupdf.Point(
                        caixa_original.x0,
                        caixa_original.y1 - 1,
                    ),
                    texto_novo,
                    fontname=fonte,
                    fontsize=tamanho_ajustado,
                    color=cor,
                    overlay=True,
                )

            conferencias.append(
                {
                    "Página": indice_pagina,
                    "Preço original": item["texto_original"],
                    "Preço atualizado": item["texto_novo"],
                    "Multiplicador": str(multiplicador),
                }
            )

    resultado = io.BytesIO()

    documento.save(
        resultado,
        garbage=4,
        deflate=True,
    )

    documento.close()
    resultado.seek(0)

    return resultado.getvalue(), conferencias


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def converter_valor_excel_para_decimal(valor):
    """
    Aceita números e textos monetários brasileiros.
    """
    if valor is None:
        return None

    if isinstance(valor, (int, float, Decimal)):
        return Decimal(str(valor))

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.upper().replace("R$", "").strip()
    texto = texto.replace(" ", "")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def parece_preco_excel(valor):
    """
    Evita modificar descrição, código, dimensão e medidas.
    """
    if isinstance(valor, (int, float, Decimal)):
        return True

    if not isinstance(valor, str):
        return False

    return bool(
        re.match(
            r"^\s*(R\$\s*)?\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\s*$",
            valor,
        )
    )


def obter_cabecalhos(planilha, linha_cabecalho):
    cabecalhos = {}

    for numero_coluna in range(1, planilha.max_column + 1):
        valor = planilha.cell(
            linha_cabecalho,
            numero_coluna,
        ).value

        if valor is not None and str(valor).strip():
            cabecalhos[str(valor).strip()] = numero_coluna

    return cabecalhos


def atualizar_excel(
    arquivo_excel,
    nome_aba,
    linha_cabecalho,
    numero_coluna,
    multiplicador,
):
    """
    Atualiza somente a coluna escolhida.
    """
    workbook = load_workbook(
        io.BytesIO(arquivo_excel),
        data_only=False,
        keep_links=True,
        rich_text=True,
    )

    planilha = workbook[nome_aba]
    conferencias = []

    for numero_linha in range(
        linha_cabecalho + 1,
        planilha.max_row + 1,
    ):
        celula = planilha.cell(
            numero_linha,
            numero_coluna,
        )

        valor_original = celula.value

        if valor_original is None:
            continue

        if isinstance(valor_original, str):
            if valor_original.startswith("="):
                continue

        if not parece_preco_excel(valor_original):
            continue

        valor_decimal = converter_valor_excel_para_decimal(
            valor_original
        )

        if valor_decimal is None:
            continue

        novo_valor = arredondar_sempre_para_cima(
            valor_decimal * Decimal(str(multiplicador))
        )

        # Se a célula original era texto com R$, preserva como texto.
        if isinstance(valor_original, str) and "R$" in valor_original.upper():
            celula.value = formatar_preco_brasileiro(novo_valor)
        else:
            celula.value = novo_valor

        conferencias.append(
            {
                "Linha": numero_linha,
                "Célula": celula.coordinate,
                "Preço original": valor_original,
                "Preço atualizado": celula.value,
            }
        )

    resultado = io.BytesIO()
    workbook.save(resultado)
    resultado.seek(0)

    return resultado.getvalue(), conferencias


# ==========================================================
# TELA DO APLICATIVO
# ==========================================================

arquivo = st.file_uploader(
    "Envie um PDF ou Excel",
    type=["pdf", "xlsx"],
)

multiplicador = st.number_input(
    "Multiplicador de preço",
    min_value=0.000001,
    value=2.383949988,
    step=0.000001,
    format="%.9f",
    help=(
        "O novo preço será o valor original multiplicado por este fator "
        "e arredondado sempre para cima."
    ),
)

if arquivo is None:
    st.info("Envie um arquivo para começar.")
    st.stop()

nome_arquivo = arquivo.name
extensao = Path(nome_arquivo).suffix.lower()
conteudo = arquivo.getvalue()


# ==========================================================
# INTERFACE DE PDF
# ==========================================================

if extensao == ".pdf":
    st.subheader("PDF → PDF")
    st.info(
        "O arquivo final será outro PDF. "
        "Fotos, imagens, páginas, descrições e a estrutura original "
        "permanecem no documento. Apenas os preços identificados no "
        "formato R$ 0,00 são substituídos."
    )

    if st.button(
        "Gerar PDF com preços atualizados",
        type="primary",
    ):
        try:
            with st.spinner(
                "Localizando preços e criando o PDF atualizado..."
            ):
                resultado, conferencias = atualizar_pdf_mantendo_layout(
                    conteudo,
                    multiplicador,
                )

            if not conferencias:
                st.warning(
                    "Nenhum preço foi identificado. "
                    "Isso pode acontecer se o PDF for escaneado como imagem "
                    "ou se os preços não estiverem no formato R$ 0,00."
                )
                st.stop()

            st.success(
                f"PDF pronto. {len(conferencias)} preço(s) foram atualizados."
            )

            st.subheader("Conferência das alterações")
            st.dataframe(
                pd.DataFrame(conferencias),
                use_container_width=True,
                hide_index=True,
            )

            nome_saida = (
                f"{Path(nome_arquivo).stem}_precos_atualizados.pdf"
            )

            st.download_button(
                label="⬇️ Baixar PDF atualizado",
                data=resultado,
                file_name=nome_saida,
                mime="application/pdf",
            )

        except Exception as erro:
            st.error(
                "Não foi possível processar o PDF. "
                f"Detalhe técnico: {erro}"
            )


# ==========================================================
# INTERFACE DE EXCEL
# ==========================================================

elif extensao == ".xlsx":
    st.subheader("Excel → Excel")

    try:
        workbook_preview = load_workbook(
            io.BytesIO(conteudo),
            data_only=False,
            keep_links=True,
            rich_text=True,
        )

        nome_aba = st.selectbox(
            "Escolha a aba que contém os preços",
            workbook_preview.sheetnames,
        )

        planilha_preview = workbook_preview[nome_aba]

        linha_cabecalho = st.number_input(
            "Número da linha dos títulos das colunas",
            min_value=1,
            max_value=max(1, planilha_preview.max_row),
            value=1,
            step=1,
        )

        cabecalhos = obter_cabecalhos(
            planilha_preview,
            linha_cabecalho,
        )

        if not cabecalhos:
            st.error(
                "Não foram encontrados títulos nessa linha. "
                "Escolha outra linha para o cabeçalho."
            )
            st.stop()

        nome_coluna = st.selectbox(
            "Escolha a coluna de preço ou valor",
            list(cabecalhos.keys()),
        )

        numero_coluna = cabecalhos[nome_coluna]

        if st.button(
            "Gerar Excel com preços atualizados",
            type="primary",
        ):
            with st.spinner(
                "Atualizando somente a coluna selecionada..."
            ):
                resultado, conferencias = atualizar_excel(
                    conteudo,
                    nome_aba,
                    linha_cabecalho,
                    numero_coluna,
                    multiplicador,
                )

            st.success(
                f"Excel pronto. {len(conferencias)} valor(es) foram atualizados."
            )

            if conferencias:
                st.subheader("Conferência das alterações")
                st.dataframe(
                    pd.DataFrame(conferencias),
                    use_container_width=True,
                    hide_index=True,
                )

            nome_saida = (
                f"{Path(nome_arquivo).stem}_precos_atualizados.xlsx"
            )

            st.download_button(
                label="⬇️ Baixar Excel atualizado",
                data=resultado,
                file_name=nome_saida,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

    except Exception as erro:
        st.error(
            "Não foi possível abrir este Excel. "
            f"Detalhe técnico: {erro}"
        )
