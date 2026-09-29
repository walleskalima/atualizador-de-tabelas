from pathlib import Path

app_code = r'''
import io
import math
import re
import tempfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_CEILING
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Atualizador de Tabelas",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Atualizador de preços — preservando o original")
st.write(
    "Excel: altera somente a coluna escolhida, preservando a estrutura "
    "do arquivo. PDF: preserva cada página visualmente e altera somente "
    "os preços detectados na coluna escolhida."
)


# ============================================================
# NÚMEROS / PREÇOS
# ============================================================

def parse_decimal(valor):
    """Converte formatos brasileiros e internacionais para Decimal."""
    if valor is None:
        return None

    if isinstance(valor, bool):
        return None

    if isinstance(valor, Decimal):
        return valor

    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return None
        return Decimal(str(valor))

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.replace("R$", "").replace("r$", "").strip()
    texto = texto.replace(" ", "")

    # 1.234,56
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")

    # 1234,56
    elif "," in texto:
        texto = texto.replace(",", ".")

    # Remove caracteres que não pertencem ao número.
    texto = re.sub(r"[^0-9.\-+]", "", texto)

    if not texto:
        return None

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def arredondar_para_cima(valor):
    """20,00 -> 20 | 20,01 -> 21 | 20,99 -> 21."""
    return int(
        Decimal(str(valor)).quantize(
            Decimal("1"),
            rounding=ROUND_CEILING,
        )
    )


def multiplicar_preco(valor, multiplicador):
    numero = parse_decimal(valor)

    if numero is None:
        return None

    return arredondar_para_cima(
        numero * multiplicador
    )


def formatar_reais(valor):
    return (
        f"R$ {int(valor):,}"
        .replace(",", ".")
    )


def regex_preco():
    """
    Aceita:
        R$ 926,82
        R$ 1.073,88
        R$1.624,00
    """
    return re.compile(
        r"R\$\s*"
        r"(?P<valor>\d{1,3}(?:\.\d{3})*(?:,\d{1,2})"
        r"|\d+(?:,\d{1,2}))",
        re.IGNORECASE,
    )


def parece_preco_texto(valor):
    if not isinstance(valor, str):
        return False

    return bool(regex_preco().search(valor))


def header_indica_preco(nome):
    """
    Ajuda a decidir se uma célula numérica é preço quando não possui R$.
    """
    if nome is None:
        return False

    texto = str(nome).strip().lower()

    termos = [
        "preço",
        "preco",
        "valor",
        "custo",
        "venda",
        "price",
        "cost",
        "amount",
        "r$",
    ]

    return any(termo in texto for termo in termos)


# ============================================================
# EXCEL — PRESERVAR O ARQUIVO ORIGINAL
# ============================================================

def abrir_excel(arquivo_bytes, nome_arquivo):
    keep_vba = nome_arquivo.lower().endswith(".xlsm")

    return load_workbook(
        io.BytesIO(arquivo_bytes),
        data_only=False,
        keep_links=True,
        keep_vba=keep_vba,
        rich_text=True,
    )


def obter_colunas_com_exemplos(planilha):
    """
    Não depende de cabeçalho em linha fixa.
    Retorna as colunas reais utilizadas e alguns exemplos.
    """
    resultado = []

    for numero_coluna in range(
        1,
        planilha.max_column + 1
    ):
        letra = get_column_letter(numero_coluna)

        valores = []

        for linha in range(
            1,
            min(planilha.max_row, 12) + 1
        ):
            valor = planilha.cell(
                linha,
                numero_coluna
            ).value

            if valor is not None:
                valores.append(str(valor))

        exemplos = " | ".join(valores[:4])

        resultado.append({
            "numero": numero_coluna,
            "letra": letra,
            "exemplos": exemplos,
        })

    return resultado


def célula_deve_ser_alterada(
    célula,
    nome_cabecalho,
    coluna_tem_apenas_precos,
):
    """
    Regra conservadora:

    - texto contendo R$ -> altera
    - número em coluna explicitamente marcada como somente preços -> altera
    - número com formato monetário -> altera
    - número em coluna cujo cabeçalho indica preço/valor -> altera
    - demais números -> NÃO altera
    """
    valor = célula.value

    if valor is None:
        return False

    if isinstance(valor, str):
        if valor.startswith("="):
            return False

        return parece_preco_texto(valor)

    if isinstance(valor, (int, float, Decimal)):
        if coluna_tem_apenas_precos:
            return True

        formato = str(
            célula.number_format or ""
        ).upper()

        if (
            "R$" in formato
            or "$" in formato
            or "BRL" in formato
        ):
            return True

        if header_indica_preco(nome_cabecalho):
            return True

    return False


def atualizar_celula_excel(célula, multiplicador):
    valor_original = célula.value

    if isinstance(valor_original, str):
        match = regex_preco().search(
            valor_original
        )

        if not match:
            return None

        valor = parse_decimal(
            match.group("valor")
        )

        if valor is None:
            return None

        novo = multiplicar_preco(
            valor,
            multiplicador
        )

        texto_novo = regex_preco().sub(
            formatar_reais(novo),
            valor_original,
        )

        return texto_novo

    if isinstance(
        valor_original,
        (int, float, Decimal)
    ):
        return multiplicar_preco(
            valor_original,
            multiplicador
        )

    return None


def atualizar_excel(
    arquivo_bytes,
    nome_arquivo,
    nome_aba,
    letra_coluna,
    multiplicador,
    coluna_tem_apenas_precos,
):
    workbook = abrir_excel(
        arquivo_bytes,
        nome_arquivo
    )

    planilha = workbook[
        nome_aba
    ]

    numero_coluna = planilha[
        f"{letra_coluna}1"
    ].column

    # Tenta identificar cabeçalho, sem exigir linha fixa.
    possiveis_cabecalhos = []
    for linha in range(
        1,
        min(planilha.max_row, 10) + 1
    ):
        valor = planilha.cell(
            linha,
            numero_coluna
        ).value

        if valor is not None:
            possiveis_cabecalhos.append(
                str(valor)
            )

    # O usuário escolheu a coluna, então usamos o melhor indício textual.
    nome_cabecalho = (
        possiveis_cabecalhos[0]
        if possiveis_cabecalhos
        else ""
    )

    alteracoes = []

    for linha in range(
        1,
        planilha.max_row + 1
    ):
        célula = planilha.cell(
            linha,
            numero_coluna
        )

        if not célula_deve_ser_alterada(
            célula,
            nome_cabecalho,
            coluna_tem_apenas_precos,
        ):
            continue

        original = célula.value

        novo = atualizar_celula_excel(
            célula,
            multiplicador
        )

        if novo is None or novo == original:
            continue

        célula.value = novo

        # Mantém o estilo existente da célula.
        # Somente quando era número e tinha formato monetário,
        # garante exibição sem casas decimais.
        if isinstance(
            novo,
            (int, float)
        ):
            formato_original = str(
                célula.number_format or ""
            )

            if (
                "R$" in formato_original.upper()
                or "$" in formato_original
            ):
                célula.number_format = (
                    'R$ #,##0'
                )

        alteracoes.append({
            "Aba": nome_aba,
            "Célula": célula.coordinate,
            "Original": original,
            "Atualizado": novo,
        })

    saida = io.BytesIO()
    workbook.save(saida)
    saida.seek(0)

    return (
        saida.getvalue(),
        alteracoes,
    )


# ============================================================
# PDF — DETECÇÃO DE PREÇOS COM COORDENADAS
# ============================================================

def eh_preco_pdf(texto):
    texto = str(texto).strip()

    return bool(
        re.fullmatch(
            r"\d{1,3}(?:\.\d{3})*,\d{2}",
            texto
        )
        or
        re.fullmatch(
            r"\d+,\d{2}",
            texto
        )
    )


def detectar_precos_pdf(pdf_bytes):
    """
    Retorna cada preço junto com:
        página
        bbox
        valor
        tamanho da fonte
        cor aproximada
        posição X
    """
    documento = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    precos = []

    for pagina_idx, pagina in enumerate(
        documento,
        start=1
    ):
        palavras = pagina.get_text(
            "words"
        )

        # Palavras no formato:
        # x0,y0,x1,y1,text,block,line,word
        for i, palavra in enumerate(
            palavras
        ):
            x0, y0, x1, y1, texto, bloco, linha, palavra_idx = palavra

            texto = str(texto).strip()

            if not eh_preco_pdf(texto):
                continue

            # Verifica se há "R$" imediatamente antes.
            tem_rs = False

            for anterior in palavras:
                ax0, ay0, ax1, ay1, atexto, ab, al, aw = anterior

                if (
                    ab == bloco
                    and al == linha
                    and abs(
                        ax1 - x0
                    ) < 40
                    and str(atexto).strip().upper()
                    == "R$"
                ):
                    tem_rs = True
                    x0_final = min(
                        x0,
                        ax0
                    )
                    break
            else:
                x0_final = x0

            if not tem_rs:
                continue

            # Tenta descobrir tamanho/cor da fonte.
            font_size = 10
            font_color = 0

            try:
                texto_dict = pagina.get_text(
                    "dict"
                )

                melhor_span = None

                for bloco_dict in texto_dict.get(
                    "blocks",
                    []
                ):
                    if bloco_dict.get(
                        "type"
                    ) != 0:
                        continue

                    for linha_dict in bloco_dict.get(
                        "lines",
                        []
                    ):
                        for span in linha_dict.get(
                            "spans",
                            []
                        ):
                            bbox = span.get(
                                "bbox"
                            )

                            if not bbox:
                                continue

                            sx0, sy0, sx1, sy1 = bbox

                            # Procura o span que contém
                            # ou intersecta o preço.
                            if (
                                sx0 <= x1
                                and sx1 >= x0
                                and sy0 <= y1
                                and sy1 >= y0
                            ):
                                melhor_span = span
                                break

                        if melhor_span:
                            break

                    if melhor_span:
                        break

                if melhor_span:
                    font_size = float(
                        melhor_span.get(
                            "size",
                            10
                        )
                    )

                    font_color = int(
                        melhor_span.get(
                            "color",
                            0
                        )
                    )

            except Exception:
                pass

            precos.append({
                "pagina": pagina_idx,
                "valor_texto": texto,
                "valor": parse_decimal(texto),
                "rect": fitz.Rect(
                    x0_final,
                    y0,
                    x1,
                    y1,
                ),
                "x_centro": (
                    x0_final + x1
                ) / 2,
                "font_size": font_size,
                "font_color": font_color,
            })

    documento.close()

    return precos


def agrupar_colunas_precos(precos, tolerancia=35):
    """
    Agrupa preços por posição horizontal.
    Isso cria uma 'coluna' mesmo que o PDF não tenha
    uma estrutura de células verdadeira.
    """
    if not precos:
        return []

    ordenados = sorted(
        precos,
        key=lambda item: item["x_centro"]
    )

    grupos = []

    for preco in ordenados:
        colocado = False

        for grupo in grupos:
            media_x = grupo["media_x"]

            if abs(
                preco["x_centro"] - media_x
            ) <= tolerancia:
                grupo["itens"].append(
                    preco
                )

                grupo["media_x"] = sum(
                    item["x_centro"]
                    for item in grupo["itens"]
                ) / len(
                    grupo["itens"]
                )

                colocado = True
                break

        if not colocado:
            grupos.append({
                "media_x": preco["x_centro"],
                "itens": [preco],
            })

    grupos.sort(
        key=lambda grupo: grupo["media_x"]
    )

    return grupos


# ============================================================
# PDF — ALTERAÇÃO VISUAL
# ============================================================

def gerar_cor_rgb_pdf(cor):
    """
    PyMuPDF retorna cor como inteiro 0xRRGGBB.
    """
    try:
        r = (cor >> 16) & 255
        g = (cor >> 8) & 255
        b = cor & 255

        return (
            r / 255,
            g / 255,
            b / 255,
        )
    except Exception:
        return (0, 0, 0)


def modificar_pdf_visualmente(
    pdf_bytes,
    precos,
    grupo_coluna,
    multiplicador,
):
    """
    Mantém o PDF inteiro original e substitui visualmente
    SOMENTE os preços da coluna selecionada.

    Depois as páginas modificadas são colocadas no Excel
    como imagens em tamanho proporcional.
    """
    documento = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    itens_alterar = set(
        id(item)
        for item in grupo_coluna["itens"]
    )

    alteracoes = []

    # 1. Redações.
    for pagina_idx, pagina in enumerate(
        documento,
        start=1
    ):
        for item in precos:
            if (
                item["pagina"] != pagina_idx
                or id(item) not in itens_alterar
            ):
                continue

            novo = multiplicar_preco(
                item["valor"],
                multiplicador
            )

            item["novo"] = novo

            # Pequena margem para apagar somente o texto.
            rect = fitz.Rect(
                item["rect"]
            )

            rect.x0 -= 1
            rect.y0 -= 1
            rect.x1 += 1
            rect.y1 += 1

            pagina.add_redact_annot(
                rect,
                fill=(1, 1, 1),
                cross_out=False,
            )

    # Aplica todas as redações.
    for pagina in documento:
        try:
            pagina.apply_redactions()
        except Exception:
            pass

    # 2. Insere os novos valores.
    for item in precos:
        if id(item) not in itens_alterar:
            continue

        pagina = documento[
            item["pagina"] - 1
        ]

        rect = fitz.Rect(
            item["rect"]
        )

        texto_novo = formatar_reais(
            item["novo"]
        )

        # Mantém tamanho aproximado e alinha à direita,
        # adequado para valores monetários.
        tamanho = max(
            6,
            min(
                item["font_size"],
                36
            )
        )

        cor = gerar_cor_rgb_pdf(
            item["font_color"]
        )

        try:
            pagina.insert_textbox(
                rect,
                texto_novo,
                fontname="helv",
                fontsize=tamanho,
                color=cor,
                align=fitz.TEXT_ALIGN_RIGHT,
                overlay=True,
            )
        except Exception:
            # Fallback para inserção simples.
            pagina.insert_text(
                (
                    rect.x1,
                    rect.y1,
                ),
                texto_novo,
                fontname="helv",
                fontsize=tamanho,
                color=cor,
                overlay=True,
            )

        alteracoes.append({
            "Página": item["pagina"],
            "Posição X": round(
                item["x_centro"],
                1
            ),
            "Original": (
                f"R$ {item['valor_texto']}"
            ),
            "Atualizado": texto_novo,
        })

    return (
        documento,
        alteracoes,
    )


# ============================================================
# PDF — COLOCAR A PÁGINA ORIGINAL NO EXCEL
# ============================================================

def inserir_pagina_como_imagem(
    ws,
    pagina,
    numero_pagina,
):
    """
    Coloca a página inteira como imagem.

    Isso preserva:
        fotos
        logos
        textos
        bordas
        proporções
        espaçamentos
        disposição original
    """
    zoom = 2.0

    pix = pagina.get_pixmap(
        matrix=fitz.Matrix(
            zoom,
            zoom
        ),
        alpha=False,
    )

    imagem = PILImage.open(
        io.BytesIO(
            pix.tobytes("png")
        )
    ).convert("RGB")

    caminho = tempfile.NamedTemporaryFile(
        suffix=f"_pagina_{numero_pagina}.png",
        delete=False,
    ).name

    imagem.save(
        caminho,
        "PNG"
    )

    imagem_excel = ExcelImage(
        caminho
    )

    # Limite apenas para não gerar um Excel absurdo.
    largura_maxima = 1400

    if imagem_excel.width > largura_maxima:
        escala = (
            largura_maxima
            / imagem_excel.width
        )

        imagem_excel.width = int(
            imagem_excel.width * escala
        )

        imagem_excel.height = int(
            imagem_excel.height * escala
        )

    ws.add_image(
        imagem_excel,
        "A1"
    )

    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = (
        max(
            20,
            imagem_excel.width / 7
        )
    )


# ============================================================
# PDF — GERAR EXCEL FINAL
# ============================================================

def gerar_excel_pdf(
    pdf_bytes,
    grupo_coluna,
    precos,
    multiplicador,
):
    documento, alteracoes = (
        modificar_pdf_visualmente(
            pdf_bytes,
            precos,
            grupo_coluna,
            multiplicador,
        )
    )

    workbook = Workbook()

    # Remove planilha inicial.
    workbook.remove(
        workbook.active
    )

    # ========================================================
    # ABA EDITÁVEL / CONFERÊNCIA
    # ========================================================

    ws_dados = workbook.create_sheet(
        "Dados editáveis",
        0
    )

    cabecalhos = [
        "Página",
        "Posição X",
        "Valor original",
        "Valor atualizado",
    ]

    for coluna, titulo in enumerate(
        cabecalhos,
        start=1
    ):
        celula = ws_dados.cell(
            1,
            coluna,
            titulo
        )

        celula.font = Font(
            bold=True,
            color="FFFFFF",
        )

        celula.fill = PatternFill(
            "solid",
            fgColor="1F4E78",
        )

        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for linha, item in enumerate(
        alteracoes,
        start=2
    ):
        ws_dados.cell(
            linha,
            1,
            item["Página"]
        )

        ws_dados.cell(
            linha,
            2,
            item["Posição X"]
        )

        ws_dados.cell(
            linha,
            3,
            item["Original"]
        )

        ws_dados.cell(
            linha,
            4,
            item["Atualizado"]
        )

    larguras = {
        "A": 12,
        "B": 14,
        "C": 22,
        "D": 22,
    }

    for coluna, largura in larguras.items():
        ws_dados.column_dimensions[
            coluna
        ].width = largura

    ws_dados.freeze_panes = "A2"

    # ========================================================
    # ABAS VISUAIS — UMA PÁGINA POR ABA
    # ========================================================

    for numero_pagina, pagina in enumerate(
        documento,
        start=1
    ):
        ws = workbook.create_sheet(
            f"Página {numero_pagina}"
        )

        inserir_pagina_como_imagem(
            ws,
            pagina,
            numero_pagina
        )

    # ========================================================
    # INFORMAÇÕES
    # ========================================================

    ws_info = workbook.create_sheet(
        "Informações"
    )

    ws_info["A1"] = (
        "Atualização de tabela PDF"
    )

    ws_info["A1"].font = Font(
        bold=True,
        size=14
    )

    ws_info["A3"] = "Multiplicador"
    ws_info["B3"] = str(
        multiplicador
    )

    ws_info["A4"] = (
        "Arredondamento"
    )

    ws_info["B4"] = (
        "Sempre para cima"
    )

    ws_info["A6"] = "Total de preços"
    ws_info["B6"] = len(
        alteracoes
    )

    ws_info["A8"] = "Preservação visual"
    ws_info["B8"] = (
        "As abas 'Página N' mantêm "
        "cada página inteira do PDF "
        "como imagem, inclusive fotos, "
        "logos, textos e disposição. "
        "Somente os preços selecionados "
        "foram substituídos visualmente."
    )

    ws_info.column_dimensions["A"].width = 28
    ws_info.column_dimensions["B"].width = 100
    ws_info["B8"].alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )

    documento.close()

    saida = io.BytesIO()

    workbook.save(saida)
    saida.seek(0)

    return (
        saida.getvalue(),
        alteracoes,
    )


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

arquivo = st.file_uploader(
    "Envie o arquivo original",
    type=[
        "xlsx",
        "xlsm",
        "pdf",
    ],
)

multiplicador_texto = st.text_input(
    "Valor a multiplicar",
    value="2,383949988",
)

multiplicador = parse_decimal(
    multiplicador_texto
)

if multiplicador is None or multiplicador <= 0:
    st.error(
        "Digite um multiplicador válido."
    )
    st.stop()

if arquivo is None:
    st.info(
        "Envie um PDF ou Excel para começar."
    )
    st.stop()

arquivo_bytes = arquivo.getvalue()
extensao = Path(
    arquivo.name
).suffix.lower()


# ============================================================
# EXCEL
# ============================================================

if extensao in [".xlsx", ".xlsm"]:

    try:
        workbook_preview = abrir_excel(
            arquivo_bytes,
            arquivo.name
        )
    except Exception as erro:
        st.error(
            f"Não foi possível abrir o Excel: {erro}"
        )
        st.stop()

    aba = st.selectbox(
        "Escolha a aba",
        workbook_preview.sheetnames,
    )

    planilha = workbook_preview[
        aba
    ]

    opcoes_colunas = (
        obter_colunas_com_exemplos(
            planilha
        )
    )

    labels = []

    for item in opcoes_colunas:
        exemplo = item["exemplos"]

        if len(exemplo) > 90:
            exemplo = exemplo[:90] + "..."

        labels.append(
            f"{item['letra']}  |  {exemplo}"
        )

    escolha = st.selectbox(
        "Escolha a coluna dos preços/valores",
        range(len(labels)),
        format_func=lambda i: labels[i],
    )

    letra_coluna = opcoes_colunas[
        escolha
    ]["letra"]

    coluna_tem_apenas_precos = st.checkbox(
        "A coluna escolhida contém somente preços/valores "
        "que devem ser multiplicados",
        value=True,
        help=(
            "Ative para multiplicar números mesmo quando "
            "a célula não contém 'R$' nem formato monetário."
        ),
    )

    st.subheader(
        "Pré-visualização da coluna"
    )

    preview = []

    for linha in range(
        1,
        min(planilha.max_row, 30) + 1
    ):
        preview.append({
            "Linha": linha,
            "Valor": planilha.cell(
                linha,
                planilha[
                    f"{letra_coluna}1"
                ].column
            ).value,
        })

    st.dataframe(
        pd.DataFrame(preview),
        use_container_width=True,
        hide_index=True,
    )

    if st.button(
        "🚀 Gerar Excel mantendo o original",
        type="primary",
    ):
        with st.spinner(
            "Atualizando somente os preços..."
        ):
            resultado, alteracoes = (
                atualizar_excel(
                    arquivo_bytes,
                    arquivo.name,
                    aba,
                    letra_coluna,
                    multiplicador,
                    coluna_tem_apenas_precos,
                )
            )

        st.success(
            f"Foram atualizados {len(alteracoes)} valor(es)."
        )

        if alteracoes:
            st.subheader(
                "Conferência"
            )

            st.dataframe(
                pd.DataFrame(
                    alteracoes
                ),
                use_container_width=True,
                hide_index=True,
            )

        extensao_saida = (
            ".xlsm"
            if extensao == ".xlsm"
            else ".xlsx"
        )

        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if extensao_saida == ".xlsm"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            "⬇️ Baixar arquivo atualizado",
            data=resultado,
            file_name=(
                f"{Path(arquivo.name).stem}"
                f"_atualizado{extensao_saida}"
            ),
            mime=mime,
        )


# ============================================================
# PDF
# ============================================================

elif extensao == ".pdf":

    st.info(
        "Neste modo o sistema NÃO reconstrói a tabela linha a linha. "
        "Ele mantém cada página do PDF inteira, como original, "
        "e troca somente os preços detectados na coluna escolhida."
    )

    with st.spinner(
        "Localizando preços e possíveis colunas..."
    ):
        precos = detectar_precos_pdf(
            arquivo_bytes
        )

    if not precos:
        st.error(
            "Nenhum preço no formato 'R$ 0,00' foi encontrado "
            "como texto no PDF."
        )

        st.warning(
            "Esse PDF pode ser escaneado como imagem. "
            "Nesse caso será necessário ativar OCR."
        )

        st.stop()

    grupos = agrupar_colunas_precos(
        precos
    )

    st.subheader(
        "Colunas de preços encontradas"
    )

    opcoes_pdf = []

    for indice, grupo in enumerate(
        grupos,
        start=1
    ):
        paginas = sorted(
            set(
                item["pagina"]
                for item in grupo["itens"]
            )
        )

        opcoes_pdf.append(
            f"Coluna {indice} | "
            f"X ≈ {grupo['media_x']:.0f} | "
            f"{len(grupo['itens'])} preços | "
            f"páginas: {', '.join(map(str, paginas[:12]))}"
        )

    coluna_pdf_idx = st.selectbox(
        "Escolha a coluna que contém os preços",
        range(len(opcoes_pdf)),
        format_func=lambda i: opcoes_pdf[i],
    )

    grupo_escolhido = grupos[
        coluna_pdf_idx
    ]

    st.subheader(
        "Preços encontrados nessa coluna"
    )

    tabela_previa = pd.DataFrame([
        {
            "Página": item["pagina"],
            "Valor original": f"R$ {item['valor_texto']}",
            "Novo valor": formatar_reais(
                multiplicar_preco(
                    item["valor"],
                    multiplicador
                )
            ),
            "X": round(
                item["x_centro"],
                1
            ),
        }
        for item in grupo_escolhido["itens"]
    ])

    st.dataframe(
        tabela_previa,
        use_container_width=True,
        hide_index=True,
    )

    if st.button(
        "🚀 Gerar Excel mantendo o visual ORIGINAL",
        type="primary",
    ):
        with st.spinner(
            "Preservando páginas e atualizando preços..."
        ):
            resultado, alteracoes = (
                gerar_excel_pdf(
                    arquivo_bytes,
                    grupo_escolhido,
                    precos,
                    multiplicador,
                )
            )

        st.success(
            f"Foram atualizados {len(alteracoes)} preço(s)."
        )

        st.caption(
            "As abas 'Página 1', 'Página 2' etc. são imagens das páginas "
            "originais com apenas os preços escolhidos substituídos. "
            "A aba 'Dados editáveis' contém a conferência dos valores."
        )

        st.dataframe(
            pd.DataFrame(
                alteracoes
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Baixar Excel",
            data=resultado,
            file_name=(
                f"{Path(arquivo.name).stem}"
                f"_atualizado.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
'''

requirements = r'''
streamlit>=1.40
openpyxl>=3.1
pandas>=2.2
PyMuPDF>=1.24
Pillow>=10.0
'''

Path("/mnt/data/app_multiplicador_tabelas_ORIGINAL.py").write_text(
    app_code,
    encoding="utf-8",
)

Path("/mnt/data/requirements_ORIGINAL.txt").write_text(
    requirements.strip() + "\n",
    encoding="utf-8",
)

print("Arquivos criados:")
print("/mnt/data/app_multiplicador_tabelas_ORIGINAL.py")
print("/mnt/data/requirements_ORIGINAL.txt")
