```python
import io
import re
import tempfile
import os
from decimal import Decimal, InvalidOperation, ROUND_CEILING
from pathlib import Path

import pandas as pd
import pymupdf
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ==========================================================
# CONFIGURAÇÃO DO APLICATIVO
# ==========================================================

st.set_page_config(
    page_title="Atualizador de Preços",
    page_icon="💰",
    layout="wide",
)

st.title("💰 Atualizador de Preços")
st.caption(
    "PDF → PDF → Excel e Excel → Excel."
)


# ==========================================================
# REGRAS DE PREÇO
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
    Exemplos:
    R$ 926,82 -> Decimal("926.82")
    R$ 1.073,88 -> Decimal("1073.88")
    """

    texto = str(texto_preco).upper()
    texto = texto.replace("R$", "").strip()
    texto = texto.replace(".", "").replace(",", ".")

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def formatar_preco_brasileiro(valor_inteiro):
    """
    Exemplos:
    2210 -> R$ 2.210
    2561 -> R$ 2.561
    """

    numero = f"{valor_inteiro:,}"
    numero = numero.replace(",", ".")

    return f"R$ {numero}"


def calcular_novo_preco(texto_preco, multiplicador):
    """
    Novo preço = valor original x multiplicador,
    com arredondamento sempre para cima.
    """

    valor_original = texto_preco_para_decimal(
        texto_preco
    )

    if valor_original is None:
        return None

    resultado = (
        valor_original
        * Decimal(str(multiplicador))
    )

    return arredondar_sempre_para_cima(
        resultado
    )


# ==========================================================
# PDF -> PDF
# ==========================================================

def rgb_inteiro_para_tupla(cor):
    """
    Converte cor RGB inteira do PyMuPDF
    para RGB com valores de 0 a 1.
    """

    if not isinstance(cor, int):
        return (0, 0, 0)

    vermelho = ((cor >> 16) & 255) / 255
    verde = ((cor >> 8) & 255) / 255
    azul = (cor & 255) / 255

    return (
        vermelho,
        verde,
        azul,
    )


def escolher_fonte_pdf(nome_fonte_original):
    """
    Seleciona uma fonte PDF segura
    para escrever o novo preço.
    """

    nome = str(nome_fonte_original).lower()

    if "bold" in nome:
        return "hebo"

    if "times" in nome:
        return "tiro"

    if "courier" in nome:
        return "cour"

    return "helv"


def medir_largura_texto(
    texto,
    fonte,
    tamanho,
):
    """
    Mede a largura do texto para ajustar
    o tamanho da fonte quando necessário.
    """

    try:
        return pymupdf.get_text_length(
            texto,
            fontname=fonte,
            fontsize=tamanho,
        )

    except Exception:
        return len(texto) * tamanho * 0.55


def localizar_precos_na_pagina(pagina):
    """
    Localiza textos como:

    R$ 926,82
    R$ 1.073,88
    """

    encontrados = []

    estrutura = pagina.get_text(
        "dict"
    )

    for bloco in estrutura.get(
        "blocks",
        [],
    ):

        if bloco.get("type") != 0:
            continue

        for linha in bloco.get(
            "lines",
            [],
        ):

            for span in linha.get(
                "spans",
                [],
            ):

                texto_span = span.get(
                    "text",
                    "",
                )

                if not texto_span:
                    continue

                correspondencias = list(
                    PADRAO_PRECO.finditer(
                        texto_span
                    )
                )

                if not correspondencias:
                    continue

                x0, y0, x1, y1 = span[
                    "bbox"
                ]

                largura_span = (
                    x1 - x0
                )

                tamanho_texto = max(
                    len(texto_span),
                    1,
                )

                for correspondencia in correspondencias:

                    inicio = (
                        correspondencia.start()
                    )

                    fim = (
                        correspondencia.end()
                    )

                    if (
                        inicio == 0
                        and fim == len(
                            texto_span
                        )
                    ):

                        caixa_preco = (
                            pymupdf.Rect(
                                x0,
                                y0,
                                x1,
                                y1,
                            )
                        )

                    else:

                        x_inicio = (
                            x0
                            + (
                                largura_span
                                * inicio
                                / tamanho_texto
                            )
                        )

                        x_fim = (
                            x0
                            + (
                                largura_span
                                * fim
                                / tamanho_texto
                            )
                        )

                        caixa_preco = (
                            pymupdf.Rect(
                                x_inicio,
                                y0,
                                x_fim,
                                y1,
                            )
                        )

                    encontrados.append(
                        {
                            "texto_original":
                                correspondencia.group(
                                    0
                                ),

                            "rect":
                                caixa_preco,

                            "fonte_original":
                                span.get(
                                    "font",
                                    "helv",
                                ),

                            "tamanho_fonte":
                                span.get(
                                    "size",
                                    10,
                                ),

                            "cor":
                                rgb_inteiro_para_tupla(
                                    span.get(
                                        "color",
                                        0,
                                    )
                                ),
                        }
                    )

    return encontrados


def colocar_fundo_branco(pagina):
    """
    Coloca branco atrás do conteúdo da página.

    overlay=False coloca o branco no fundo,
    sem cobrir textos, fotos ou logos
    já existentes no PDF.
    """

    pagina.draw_rect(
        pagina.rect,
        color=None,
        fill=(1, 1, 1),
        overlay=False,
    )


def atualizar_pdf_mantendo_layout(
    arquivo_pdf,
    multiplicador,
    deixar_fundo_branco=True,
):
    """
    Atualiza preços diretamente no PDF.

    O resultado continua sendo um PDF com
    páginas, fotos, logos e estrutura
    original preservados.
    """

    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    conferencias = []

    if deixar_fundo_branco:

        for pagina in documento:
            colocar_fundo_branco(
                pagina
            )

    for (
        numero_pagina,
        pagina,
    ) in enumerate(
        documento,
        start=1,
    ):

        precos_pagina = (
            localizar_precos_na_pagina(
                pagina
            )
        )

        substituicoes = []

        for preco in precos_pagina:

            texto_original = (
                preco["texto_original"]
            )

            novo_valor = (
                calcular_novo_preco(
                    texto_original,
                    multiplicador,
                )
            )

            if novo_valor is None:
                continue

            texto_novo = (
                formatar_preco_brasileiro(
                    novo_valor
                )
            )

            caixa_original = (
                preco["rect"]
            )

            caixa_redacao = (
                pymupdf.Rect(
                    caixa_original.x0 - 1.5,
                    caixa_original.y0 - 1.0,
                    caixa_original.x1 + 1.5,
                    caixa_original.y1 + 1.0,
                )
            )

            pagina.add_redact_annot(
                caixa_redacao,
                fill=(1, 1, 1),
            )

            substituicoes.append(
                {
                    "texto_original":
                        texto_original,

                    "texto_novo":
                        texto_novo,

                    "caixa_original":
                        caixa_original,

                    "fonte":
                        escolher_fonte_pdf(
                            preco[
                                "fonte_original"
                            ]
                        ),

                    "tamanho":
                        preco[
                            "tamanho_fonte"
                        ],

                    "cor":
                        preco["cor"],
                }
            )

        if substituicoes:

            pagina.apply_redactions(
                images=(
                    pymupdf
                    .PDF_REDACT_IMAGE_NONE
                ),

                graphics=(
                    pymupdf
                    .PDF_REDACT_LINE_ART_NONE
                ),

                text=(
                    pymupdf
                    .PDF_REDACT_TEXT_REMOVE
                ),
            )

        for item in substituicoes:

            caixa = item[
                "caixa_original"
            ]

            texto_novo = item[
                "texto_novo"
            ]

            fonte = item[
                "fonte"
            ]

            cor = item[
                "cor"
            ]

            tamanho_ajustado = item[
                "tamanho"
            ]

            largura_original = (
                caixa.width
            )

            largura_nova = (
                medir_largura_texto(
                    texto_novo,
                    fonte,
                    tamanho_ajustado,
                )
            )

            while (
                largura_nova
                > largura_original
                and tamanho_ajustado > 5
            ):

                tamanho_ajustado -= 0.2

                largura_nova = (
                    medir_largura_texto(
                        texto_novo,
                        fonte,
                        tamanho_ajustado,
                    )
                )

            caixa_texto = (
                pymupdf.Rect(
                    caixa.x0,
                    caixa.y0 - 0.5,
                    max(
                        caixa.x1,
                        caixa.x0
                        + largura_nova
                        + 3,
                    ),
                    caixa.y1 + 2,
                )
            )

            resultado_texto = (
                pagina.insert_textbox(
                    caixa_texto,
                    texto_novo,
                    fontname=fonte,
                    fontsize=tamanho_ajustado,
                    color=cor,
                    align=(
                        pymupdf
                        .TEXT_ALIGN_LEFT
                    ),
                    overlay=True,
                )
            )

            if resultado_texto < 0:

                pagina.insert_text(
                    pymupdf.Point(
                        caixa.x0,
                        caixa.y1 - 1,
                    ),

                    texto_novo,

                    fontname=fonte,
                    fontsize=tamanho_ajustado,
                    color=cor,
                    overlay=True,
                )

            conferencias.append(
                {
                    "Página":
                        numero_pagina,

                    "Preço original":
                        item[
                            "texto_original"
                        ],

                    "Preço atualizado":
                        item[
                            "texto_novo"
                        ],
                }
            )

    arquivo_final = io.BytesIO()

    documento.save(
        arquivo_final,
        garbage=4,
        deflate=True,
    )

    documento.close()

    arquivo_final.seek(0)

    return (
        arquivo_final.getvalue(),
        conferencias,
    )


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def converter_valor_excel_para_decimal(
    valor
):
    """
    Converte valor numérico ou texto
    no formato brasileiro para Decimal.
    """

    if valor is None:
        return None

    if isinstance(
        valor,
        (
            int,
            float,
            Decimal,
        ),
    ):

        return Decimal(
            str(valor)
        )

    texto = str(valor).strip()

    if not texto:
        return None

    texto = (
        texto
        .upper()
        .replace("R$", "")
        .strip()
    )

    texto = texto.replace(
        " ",
        "",
    )

    if (
        "," in texto
        and "." in texto
    ):

        texto = (
            texto
            .replace(".", "")
            .replace(",", ".")
        )

    elif "," in texto:

        texto = texto.replace(
            ",",
            ".",
        )

    try:
        return Decimal(texto)

    except InvalidOperation:
        return None


def parece_preco_excel(valor):
    """
    Evita alterar medidas, códigos
    e descrições.
    """

    if isinstance(
        valor,
        (
            int,
            float,
            Decimal,
        ),
    ):

        return True

    if not isinstance(
        valor,
        str,
    ):

        return False

    padrao = (
        r"^\s*"
        r"(R\$\s*)?"
        r"\d{1,3}(?:\.\d{3})*"
        r"(?:,\d{1,2})?"
        r"\s*$"
    )

    return bool(
        re.match(
            padrao,
            valor,
        )
    )


def obter_cabecalhos(
    planilha,
    linha_cabecalho,
):
    """
    Lê títulos de colunas
    na linha escolhida.
    """

    cabecalhos = {}

    for numero_coluna in range(
        1,
        planilha.max_column + 1,
    ):

        valor = planilha.cell(
            linha_cabecalho,
            numero_coluna,
        ).value

        if (
            valor is None
            or not str(valor).strip()
        ):

            continue

        nome = str(
            valor
        ).strip()

        if nome in cabecalhos:

            nome = (
                f"{nome} "
                f"({numero_coluna})"
            )

        cabecalhos[
            nome
        ] = numero_coluna

    return cabecalhos


def atualizar_excel(
    arquivo_excel,
    nome_aba,
    linha_cabecalho,
    numero_coluna,
    multiplicador,
):
    """
    Atualiza apenas a coluna selecionada.
    """

    workbook = load_workbook(
        io.BytesIO(arquivo_excel),
        data_only=False,
        keep_links=True,
        rich_text=True,
    )

    planilha = workbook[
        nome_aba
    ]

    conferencias = []

    for numero_linha in range(
        linha_cabecalho + 1,
        planilha.max_row + 1,
    ):

        celula = planilha.cell(
            numero_linha,
            numero_coluna,
        )

        valor_original = (
            celula.value
        )

        if valor_original is None:
            continue

        if (
            isinstance(
                valor_original,
                str,
            )
            and valor_original.startswith(
                "="
            )
        ):

            continue

        if not parece_preco_excel(
            valor_original
        ):

            continue

        valor_decimal = (
            converter_valor_excel_para_decimal(
                valor_original
            )
        )

        if valor_decimal is None:
            continue

        novo_valor = (
            arredondar_sempre_para_cima(
                valor_decimal
                * Decimal(
                    str(
                        multiplicador
                    )
                )
            )
        )

        if (
            isinstance(
                valor_original,
                str,
            )
            and "R$"
            in valor_original.upper()
        ):

            celula.value = (
                formatar_preco_brasileiro(
                    novo_valor
                )
            )

        else:

            celula.value = novo_valor

            formato_anterior = str(
                celula.number_format
            )

            if (
                "R$"
                in formato_anterior
                or "$"
                in formato_anterior
            ):

                celula.number_format = (
                    "R$ #,##0"
                )

        conferencias.append(
            {
                "Linha":
                    numero_linha,

                "Célula":
                    celula.coordinate,

                "Valor original":
                    valor_original,

                "Valor atualizado":
                    celula.value,
            }
        )

    arquivo_final = io.BytesIO()

    workbook.save(
        arquivo_final
    )

    arquivo_final.seek(0)

    return (
        arquivo_final.getvalue(),
        conferencias,
    )


# ==========================================================
# PDF -> EXCEL
# ==========================================================

def inserir_pagina_pdf_como_imagem(
    planilha,
    pagina,
):
    """
    Coloca uma página inteira do PDF
    dentro do Excel como imagem.

    Isso preserva visualmente:

    - colunas
    - textos
    - imagens
    - logotipos
    - cores
    - fundos
    - posições
    - espaçamentos
    - proporções
    """

    # ------------------------------------------------------
    # RENDERIZAÇÃO EM ALTA RESOLUÇÃO
    # ------------------------------------------------------

    escala = 2.5

    pix = pagina.get_pixmap(
        matrix=pymupdf.Matrix(
            escala,
            escala,
        ),
        alpha=False,
    )

    imagem_bytes = pix.tobytes(
        "png"
    )

    # ------------------------------------------------------
    # ARQUIVO TEMPORÁRIO
    # ------------------------------------------------------

    arquivo_temporario = (
        tempfile.NamedTemporaryFile(
            suffix=".png",
            delete=False,
        )
    )

    try:

        arquivo_temporario.write(
            imagem_bytes
        )

        arquivo_temporario.close()

        # --------------------------------------------------
        # CRIA IMAGEM DO EXCEL
        # --------------------------------------------------

        imagem_excel = ExcelImage(
            arquivo_temporario.name
        )

        # Mantém exatamente a proporção
        # renderizada da página.
        imagem_excel.width = pix.width
        imagem_excel.height = pix.height

        # Coloca no canto superior esquerdo.
        planilha.add_image(
            imagem_excel,
            "A1",
        )

        # --------------------------------------------------
        # EXCEL SEM GRADE
        # --------------------------------------------------

        planilha.sheet_view.showGridLines = (
            False
        )

        # --------------------------------------------------
        # MARGENS
        # --------------------------------------------------

        planilha.page_margins.left = 0
        planilha.page_margins.right = 0
        planilha.page_margins.top = 0
        planilha.page_margins.bottom = 0
        planilha.page_margins.header = 0
        planilha.page_margins.footer = 0

        # --------------------------------------------------
        # CONFIGURAÇÃO DE IMPRESSÃO
        # --------------------------------------------------

        largura_pdf = pagina.rect.width
        altura_pdf = pagina.rect.height

        if largura_pdf > altura_pdf:

            planilha.page_setup.orientation = (
                "landscape"
            )

        else:

            planilha.page_setup.orientation = (
                "portrait"
            )

        planilha.page_setup.fitToWidth = 1
        planilha.page_setup.fitToHeight = 1

        planilha.sheet_properties.pageSetUpPr.fitToPage = (
            True
        )

        # A área de impressão é mantida
        # na região da página.
        planilha.print_area = (
            "A1:A1"
        )

        # --------------------------------------------------
        # CONFIGURAÇÃO DA COLUNA/LINHA
        # --------------------------------------------------

        planilha.column_dimensions[
            "A"
        ].width = 2

        planilha.row_dimensions[
            1
        ].height = 2

    finally:

        # O Excel já carregou a imagem.
        # Podemos apagar o arquivo temporário.
        try:
            os.unlink(
                arquivo_temporario.name
            )
        except Exception:
            pass


def converter_pdf_para_excel(
    arquivo_pdf,
):
    """
    Converte o PDF para Excel mantendo
    a aparência visual das páginas.

    Cada página do PDF vira uma aba
    do Excel.

    Não cria tabelas artificiais,
    cabeçalhos artificiais ou abas
    de texto extraído.
    """

    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    workbook = Workbook()

    # Remove a aba padrão.
    primeira_aba = (
        workbook.active
    )

    workbook.remove(
        primeira_aba
    )

    quantidade_paginas = len(
        documento
    )

    for (
        numero_pagina,
        pagina,
    ) in enumerate(
        documento,
        start=1,
    ):

        # --------------------------------------------------
        # CRIA UMA ABA PARA CADA PÁGINA
        # --------------------------------------------------

        planilha = workbook.create_sheet(
            title=f"Página {numero_pagina}"
        )

        # --------------------------------------------------
        # INSERE A PÁGINA INTEIRA
        # --------------------------------------------------

        inserir_pagina_pdf_como_imagem(
            planilha,
            pagina,
        )

    documento.close()

    # ------------------------------------------------------
    # SALVA EXCEL
    # ------------------------------------------------------

    arquivo_final = io.BytesIO()

    workbook.save(
        arquivo_final
    )

    arquivo_final.seek(0)

    return (
        arquivo_final.getvalue(),
        quantidade_paginas,
    )


# ==========================================================
# INTERFACE
# ==========================================================

arquivo = st.file_uploader(
    "Envie um arquivo PDF ou Excel",
    type=[
        "pdf",
        "xlsx",
    ],
)

multiplicador = st.number_input(
    "Multiplicador de preço",
    min_value=0.000001,
    value=2.383949988,
    step=0.000001,
    format="%.9f",
)

if arquivo is None:

    st.info(
        "Envie um arquivo para começar."
    )

    st.stop()


nome_arquivo = arquivo.name

extensao = (
    Path(nome_arquivo)
    .suffix
    .lower()
)

conteudo_arquivo = (
    arquivo.getvalue()
)


# ==========================================================
# TELA: PDF
# ==========================================================

if extensao == ".pdf":

    st.subheader("PDF")

    st.info(
        "O PDF será atualizado primeiro. "
        "Depois você poderá baixar o PDF e, "
        "se quiser, converter o PDF atualizado "
        "para Excel mantendo o layout visual."
    )

    deixar_fundo_branco = st.checkbox(
        "Adicionar fundo branco atrás dos elementos",
        value=True,
        help=(
            "A camada branca entra atrás do conteúdo. "
            "Imagens de fundo existentes podem continuar visíveis."
        ),
    )

    st.info(
        "O sistema procura preços no formato "
        "R$ 0,00 e altera somente esses valores."
    )

    if st.button(
        "Gerar PDF com preços atualizados",
        type="primary",
    ):

        try:

            with st.spinner(
                "Atualizando preços no PDF..."
            ):

                resultado, conferencias = (
                    atualizar_pdf_mantendo_layout(
                        arquivo_pdf=(
                            conteudo_arquivo
                        ),

                        multiplicador=(
                            multiplicador
                        ),

                        deixar_fundo_branco=(
                            deixar_fundo_branco
                        ),
                    )
                )

            if not conferencias:

                st.warning(
                    "Nenhum preço foi identificado. "
                    "O PDF pode ser escaneado ou usar "
                    "outro formato."
                )

                st.stop()

            st.success(
                "PDF pronto. "
                f"{len(conferencias)} "
                "preço(s) atualizado(s)."
            )

            # --------------------------------------------------
            # CONFERÊNCIA
            # --------------------------------------------------

            st.subheader(
                "Conferência"
            )

            st.dataframe(
                pd.DataFrame(
                    conferencias
                ),

                use_container_width=True,

                hide_index=True,
            )

            # --------------------------------------------------
            # DOWNLOAD PDF
            # --------------------------------------------------

            nome_saida_pdf = (
                f"{Path(nome_arquivo).stem}"
                f"_precos_atualizados.pdf"
            )

            st.download_button(
                label=(
                    "⬇️ Baixar PDF atualizado"
                ),

                data=resultado,

                file_name=(
                    nome_saida_pdf
                ),

                mime="application/pdf",
            )

            # --------------------------------------------------
            # PDF ATUALIZADO -> EXCEL
            # --------------------------------------------------

            st.divider()

            st.subheader(
                "PDF atualizado → Excel"
            )

            st.info(
                "O PDF atualizado está pronto. "
                "A conversão abaixo utiliza o PDF "
                "já modificado e mantém cada página "
                "como uma cópia visual no Excel."
            )

            if st.button(
                "📊 Converter PDF atualizado para Excel",
                type="secondary",
            ):

                try:

                    with st.spinner(
                        "Convertendo PDF atualizado para Excel..."
                    ):

                        (
                            excel_resultado,
                            quantidade_paginas,
                        ) = (
                            converter_pdf_para_excel(
                                resultado
                            )
                        )

                    st.success(
                        "Excel pronto. "
                        f"{quantidade_paginas} "
                        "página(s) foram convertidas."
                    )

                    # --------------------------------------------------
                    # NOME DO EXCEL
                    # --------------------------------------------------

                    nome_saida_excel = (
                        f"{Path(nome_arquivo).stem}"
                        f"_precos_atualizados.xlsx"
                    )

                    st.download_button(
                        label=(
                            "⬇️ Baixar Excel"
                        ),

                        data=(
                            excel_resultado
                        ),

                        file_name=(
                            nome_saida_excel
                        ),

                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),

                        key=(
                            "baixar_excel_pdf_atualizado"
                        ),
                    )

                except Exception as erro:

                    st.error(
                        "Não foi possível converter "
                        "o PDF atualizado para Excel. "
                        "Detalhe técnico: "
                        + str(erro)
                    )


# ==========================================================
# TELA: EXCEL
# ==========================================================

elif extensao == ".xlsx":

    st.subheader(
        "Excel → Excel"
    )

    try:

        workbook_preview = (
            load_workbook(
                io.BytesIO(
                    conteudo_arquivo
                ),

                data_only=False,

                keep_links=True,

                rich_text=True,
            )
        )

        nome_aba = st.selectbox(
            "1. Escolha a aba que contém a tabela",

            workbook_preview.sheetnames,
        )

        planilha_preview = (
            workbook_preview[
                nome_aba
            ]
        )

        linha_cabecalho = (
            st.number_input(
                "2. Em qual linha estão os títulos das colunas?",

                min_value=1,

                max_value=max(
                    1,
                    planilha_preview.max_row,
                ),

                value=1,

                step=1,
            )
        )

        cabecalhos = (
            obter_cabecalhos(
                planilha_preview,
                linha_cabecalho,
            )
        )

        if not cabecalhos:

            st.error(
                "Não encontrei títulos nessa linha. "
                "Escolha outro número para a linha de cabeçalho."
            )

            st.stop()

        nome_coluna = (
            st.selectbox(
                "3. Escolha a coluna de preços/valores",

                list(
                    cabecalhos.keys()
                ),
            )
        )

        numero_coluna = (
            cabecalhos[
                nome_coluna
            ]
        )

        st.caption(
            f"Coluna selecionada: "
            f"{nome_coluna} "
            f"({get_column_letter(numero_coluna)})."
        )

        # ------------------------------------------------------
        # PRÉVIA
        # ------------------------------------------------------

        dados_previa = []

        for numero_linha in range(
            linha_cabecalho + 1,

            min(
                planilha_preview.max_row + 1,

                linha_cabecalho + 11,
            ),
        ):

            valor_original = (
                planilha_preview.cell(
                    numero_linha,
                    numero_coluna,
                ).value
            )

            if valor_original is None:
                continue

            valor_decimal = (
                converter_valor_excel_para_decimal(
                    valor_original
                )
            )

            if valor_decimal is None:

                novo_valor = (
                    "Não será alterado"
                )

            else:

                novo_valor = (
                    arredondar_sempre_para_cima(
                        valor_decimal
                        * Decimal(
                            str(
                                multiplicador
                            )
                        )
                    )
                )

            dados_previa.append(
                {
                    "Linha":
                        numero_linha,

                    "Valor original":
                        valor_original,

                    "Novo valor":
                        novo_valor,
                }
            )

        if dados_previa:

            st.subheader(
                "Prévia"
            )

            st.dataframe(
                pd.DataFrame(
                    dados_previa
                ),

                use_container_width=True,

                hide_index=True,
            )

        # ------------------------------------------------------
        # GERAR EXCEL
        # ------------------------------------------------------

        if st.button(
            "Gerar Excel atualizado",
            type="primary",
        ):

            with st.spinner(
                "Atualizando valores da coluna selecionada..."
            ):

                (
                    resultado,
                    conferencias,
                ) = atualizar_excel(
                    arquivo_excel=(
                        conteudo_arquivo
                    ),

                    nome_aba=(
                        nome_aba
                    ),

                    linha_cabecalho=(
                        linha_cabecalho
                    ),

                    numero_coluna=(
                        numero_coluna
                    ),

                    multiplicador=(
                        multiplicador
                    ),
                )

            if not conferencias:

                st.warning(
                    "Nenhum valor foi alterado. "
                    "Verifique se a coluna selecionada "
                    "contém preços ou números."
                )

                st.stop()

            st.success(
                f"Excel pronto. "
                f"{len(conferencias)} "
                "valor(es) atualizado(s)."
            )

            # --------------------------------------------------
            # CONFERÊNCIA
            # --------------------------------------------------

            st.subheader(
                "Conferência"
            )

            st.dataframe(
                pd.DataFrame(
                    conferencias
                ),

                use_container_width=True,

                hide_index=True,
            )

            # --------------------------------------------------
            # DOWNLOAD
            # --------------------------------------------------

            nome_saida = (
                f"{Path(nome_arquivo).stem}"
                f"_precos_atualizados.xlsx"
            )

            st.download_button(
                label=(
                    "⬇️ Baixar Excel atualizado"
                ),

                data=resultado,

                file_name=(
                    nome_saida
                ),

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

    except Exception as erro:

        st.error(
            "Não foi possível abrir ou salvar este Excel. "
            "Detalhe técnico: "
            + str(erro)
        )
```

### O que foi alterado

A parte de **Excel → Excel permanece com a mesma lógica** do seu código original. A alteração principal foi no PDF: anteriormente o programa tentava extrair tabelas, texto e criar uma estrutura nova de Excel.

Agora o fluxo é:

**PDF original → multiplicador → PDF atualizado → botão para converter → Excel com cada página do PDF como imagem.**

Isso também significa que o botão de Excel **só aparece depois que o PDF foi processado**, exatamente como você pediu.

Um detalhe: essa versão prioriza **fidelidade visual**. Portanto, no Excel a página é uma imagem, e não uma tabela reconstruída em células editáveis. É justamente isso que permite preservar o layout original do PDF com muito mais precisão.
