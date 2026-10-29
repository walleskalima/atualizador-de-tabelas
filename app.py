import io
import re
import tempfile
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
# CONFIGURAÇÃO
# ==========================================================

st.set_page_config(
    page_title="Atualizador de Preços",
    page_icon="💰",
    layout="wide",
)

st.title("💰 Atualizador de Preços")
st.caption("PDF → PDF, PDF → Excel e Excel → Excel.")


# ==========================================================
# PREÇOS E ARREDONDAMENTO
# ==========================================================

PADRAO_PRECO = re.compile(
    r"(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}"
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
    """Converte 2210 em R$ 2.210."""
    numero = f"{valor_inteiro:,}"
    numero = numero.replace(",", ".")
    return f"R$ {numero}"


def calcular_novo_preco(texto_preco, multiplicador):
    """Novo preço = teto(preço original x multiplicador)."""
    valor_original = texto_preco_para_decimal(texto_preco)

    if valor_original is None:
        return None

    resultado = valor_original * Decimal(str(multiplicador))
    return arredondar_sempre_para_cima(resultado)


# ==========================================================
# PDF -> PDF
# ==========================================================

def rgb_inteiro_para_tupla(cor):
    if not isinstance(cor, int):
        return (0, 0, 0)

    vermelho = ((cor >> 16) & 255) / 255
    verde = ((cor >> 8) & 255) / 255
    azul = (cor & 255) / 255

    return (vermelho, verde, azul)


def escolher_fonte_pdf(nome_fonte_original):
    nome = str(nome_fonte_original).lower()

    if "bold" in nome:
        return "hebo"
    if "times" in nome:
        return "tiro"
    if "courier" in nome:
        return "cour"

    return "helv"


def medir_largura_texto(texto, fonte, tamanho):
    try:
        return pymupdf.get_text_length(
            texto,
            fontname=fonte,
            fontsize=tamanho,
        )
    except Exception:
        return len(texto) * tamanho * 0.55


def localizar_cabecalhos_tabelas(pagina):
    """Retorna o centro horizontal dos cabeçalhos TAB 1, TAB 2 e TAB 3."""
    centros = {}
    estrutura = pagina.get_text("dict")

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue

        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                texto = span.get("text", "").strip()
                correspondencia = re.fullmatch(
                    r"TAB\s*([123])", texto, re.IGNORECASE
                )

                if correspondencia:
                    numero = int(correspondencia.group(1))
                    x0, _, x1, _ = span["bbox"]
                    centros[numero] = (x0 + x1) / 2

    return centros


def localizar_precos_na_pagina(pagina, colunas_precos=(1,)):
    """Localiza preços pelas coordenadas das colunas TAB do PDF."""
    encontrados = []
    centros = localizar_cabecalhos_tabelas(pagina)

    if not centros:
        return encontrados

    estrutura = pagina.get_text("dict")
    centros_ordenados = sorted(
        centros.items(),
        key=lambda item: item[1],
    )

    # Calcula os limites de cada coluna com base nos cabeçalhos.
    limites = {}

    for indice, (numero, centro) in enumerate(centros_ordenados):
        if indice > 0:
            esquerda = (
                centros_ordenados[indice - 1][1] + centro
            ) / 2
        elif len(centros_ordenados) > 1:
            distancia_direita = (
                centros_ordenados[1][1] - centro
            )
            esquerda = centro - distancia_direita / 2
        else:
            esquerda = centro - 25

        if indice < len(centros_ordenados) - 1:
            direita = (
                centro + centros_ordenados[indice + 1][1]
            ) / 2
        elif len(centros_ordenados) > 1:
            distancia_esquerda = (
                centro - centros_ordenados[indice - 1][1]
            )
            direita = centro + distancia_esquerda / 2
        else:
            direita = centro + 25

        limites[numero] = (
            max(0, esquerda),
            min(pagina.rect.width, direita),
        )

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue

        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                texto_span = span.get("text", "")

                if not texto_span:
                    continue

                for correspondencia in PADRAO_PRECO.finditer(
                    texto_span
                ):
                    x0, y0, x1, y1 = span["bbox"]
                    largura_span = x1 - x0
                    tamanho_span = max(len(texto_span), 1)
                    inicio, fim = correspondencia.span()

                    if inicio == 0 and fim == len(texto_span):
                        caixa_preco = pymupdf.Rect(
                            x0, y0, x1, y1
                        )
                    else:
                        caixa_preco = pymupdf.Rect(
                            x0 + largura_span * inicio / tamanho_span,
                            y0,
                            x0 + largura_span * fim / tamanho_span,
                            y1,
                        )

                    centro_preco = (
                        caixa_preco.x0 + caixa_preco.x1
                    ) / 2

                    coluna_encontrada = None

                    for numero, (esquerda, direita) in limites.items():
                        if esquerda <= centro_preco <= direita:
                            coluna_encontrada = numero
                            break

                    if coluna_encontrada not in colunas_precos:
                        continue

                    encontrados.append({
                        "coluna": coluna_encontrada,
                        "texto_original": correspondencia.group(0),
                        "rect": caixa_preco,
                        "fonte_original": span.get("font", "helv"),
                        "tamanho_fonte": span.get("size", 10),
                        "cor": rgb_inteiro_para_tupla(
                            span.get("color", 0)
                        ),
                    })

    return encontrados

def remover_colunas_pdf(pagina, colunas_remover):
    """Apaga cabeçalhos e preços das colunas escolhidas."""
    colunas_remover = set(colunas_remover or [])
    if not colunas_remover:
        return

    centros = localizar_cabecalhos_tabelas(pagina)
    if not centros:
        return

    estrutura = pagina.get_text("dict")
    caixas_remover = []

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue

        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                texto = span.get("text", "").strip()
                m = re.fullmatch(
                    r"TAB\s*([123])", texto, re.IGNORECASE
                )
                if m and int(m.group(1)) in colunas_remover:
                    caixas_remover.append(pymupdf.Rect(span["bbox"]))

    for preco in localizar_precos_na_pagina(
        pagina, tuple(colunas_remover)
    ):
        caixas_remover.append(preco["rect"])

    for caixa in caixas_remover:
        pagina.add_redact_annot(
            pymupdf.Rect(
                caixa.x0 - 1,
                caixa.y0 - 0.5,
                caixa.x1 + 1,
                caixa.y1 + 0.5,
            ),
            fill=(1, 1, 1),
        )

    if caixas_remover:
        pagina.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )

    colunas_mantidas = [
        n for n in centros if n not in colunas_remover
    ]

    if colunas_mantidas:
        ultima_mantida = max(
            colunas_mantidas, key=lambda n: centros[n]
        )

        if all(
            n in colunas_remover
            for n in centros
            if centros[n] > centros[ultima_mantida]
        ):
            limite_direito = min(
                pagina.rect.width, centros[ultima_mantida] + 50
            )
            if limite_direito > 100:
                pagina.set_cropbox(
                    pymupdf.Rect(
                        0, 0, limite_direito, pagina.rect.height
                    )
                )


def colocar_fundo_branco(pagina):
    pagina.draw_rect(
        pagina.rect,
        color=None,
        fill=(1, 1, 1),
        overlay=False,
    )


def ampliar_catalogo_para_impressao(pdf_bytes, grupos_por_pagina=4):
    """
    Reorganiza as páginas originais em folhas A4 horizontais.
    Divide o conteúdo por faixas para acomodar vários produtos
    por página, preservando a proporção do conteúdo original.
    """
    origem = pymupdf.open(
        stream=pdf_bytes,
        filetype="pdf",
    )
    saida = pymupdf.open()

    largura_pagina = 842
    altura_pagina = 595
    margem = 18

    for pagina_origem in origem:
        largura = pagina_origem.rect.width
        altura = pagina_origem.rect.height
        separadores = []

        for desenho in pagina_origem.get_drawings():
            for item in desenho.get("items", []):
                if item[0] == "l":
                    a, b = item[1], item[2]
                    if (
                        abs(a.y - b.y) <= 1.5
                        and abs(a.x - b.x) >= largura * 0.72
                    ):
                        y = (a.y + b.y) / 2
                        if altura * 0.30 < y < altura * 0.72:
                            separadores.append(y)

        if separadores:
            corte = min(
                separadores,
                key=lambda y: abs(y - altura / 2),
            )
        else:
            corte = altura / 2

        faixas = [(0.0, corte), (corte, altura)]

        for y0, y1 in faixas:
            clip = pymupdf.Rect(
                0, max(0, y0), largura, min(altura, y1)
            )

            if clip.height < 80:
                continue

            texto_faixa = pagina_origem.get_text(
                "text", clip=clip
            ).strip()

            if not texto_faixa:
                continue

            pagina_nova = saida.new_page(
                width=largura_pagina,
                height=altura_pagina,
            )

            escala = min(
                (largura_pagina - 2 * margem) / clip.width,
                (altura_pagina - 2 * margem) / clip.height,
            )

            w = clip.width * escala
            h = clip.height * escala
            x0_destino = (largura_pagina - w) / 2
            y0_destino = (altura_pagina - h) / 2

            destino = pymupdf.Rect(
                x0_destino,
                y0_destino,
                x0_destino + w,
                y0_destino + h,
            )

            pagina_nova.show_pdf_page(
                destino,
                origem,
                pagina_origem.number,
                clip=clip,
                keep_proportion=True,
            )

    resultado = saida.tobytes(
        garbage=4,
        deflate=True,
    )
    saida.close()
    origem.close()

    return resultado

def atualizar_pdf_mantendo_layout(
    arquivo_pdf,
    multiplicador,
    deixar_fundo_branco=True,
    colunas_remover=(),
    modo_horizontal=False,
):
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    conferencias = []

    for numero_pagina, pagina in enumerate(documento, start=1):
        precos_pagina = localizar_precos_na_pagina(
            pagina,
            colunas_precos=(1,),
        )

        substituicoes = []

        for preco in precos_pagina:
            texto_original = preco["texto_original"]
            novo_valor = calcular_novo_preco(
                texto_original,
                multiplicador,
            )

            if novo_valor is None:
                continue

            texto_novo = formatar_preco_brasileiro(novo_valor)
            caixa = preco["rect"]

            pagina.add_redact_annot(
                pymupdf.Rect(
                    caixa.x0 - 1.5,
                    caixa.y0 - 1,
                    caixa.x1 + 1.5,
                    caixa.y1 + 1,
                ),
                fill=(1, 1, 1),
            )

            substituicoes.append({
                "texto_original": texto_original,
                "texto_novo": texto_novo,
                "caixa": caixa,
                "fonte": escolher_fonte_pdf(
                    preco["fonte_original"]
                ),
                "tamanho": preco["tamanho_fonte"],
                "cor": preco["cor"],
            })

        if substituicoes:
            pagina.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,
                graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                text=pymupdf.PDF_REDACT_TEXT_REMOVE,
            )

        for item in substituicoes:
            caixa = item["caixa"]
            fonte = item["fonte"]
            tamanho = item["tamanho"]
            texto = item["texto_novo"]

            largura_original = caixa.width
            largura_nova = medir_largura_texto(
                texto,
                fonte,
                tamanho,
            )

            while largura_nova > largura_original and tamanho > 5:
                tamanho -= 0.2
                largura_nova = medir_largura_texto(
                    texto,
                    fonte,
                    tamanho,
                )

            pagina.insert_text(
                pymupdf.Point(caixa.x0, caixa.y1 - 1),
                texto,
                fontname=fonte,
                fontsize=tamanho,
                color=item["cor"],
                overlay=True,
            )

            conferencias.append({
                "Página": numero_pagina,
                "Preço original": item["texto_original"],
                "Preço atualizado": texto,
            })

        remover_colunas_pdf(pagina, colunas_remover)

    resultado = io.BytesIO()
    documento.save(resultado, garbage=4, deflate=True)
    documento.close()

    conteudo = resultado.getvalue()

    if modo_horizontal:
        conteudo = ampliar_catalogo_para_impressao(conteudo)

    return conteudo, conferencias


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def converter_valor_excel_para_decimal(valor):
    if valor is None:
        return None

    if isinstance(valor, (int, float, Decimal)):
        return Decimal(str(valor))

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.upper().replace("R$", "").replace(" ", "")

    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def parece_preco_excel(valor):
    if isinstance(valor, (int, float, Decimal)):
        return True

    if not isinstance(valor, str):
        return False

    padrao = (
        r"^\s*(?:R\$\s*)?"
        r"\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?\s*$"
    )

    return bool(re.fullmatch(padrao, valor))


def atualizar_planilha_excel(
    arquivo_excel,
    multiplicador,
    nome_aba=None,
):
    arquivo = load_workbook(io.BytesIO(arquivo_excel))

    if nome_aba and nome_aba in arquivo.sheetnames:
        abas = [arquivo[nome_aba]]
    else:
        abas = arquivo.worksheets

    abas_processadas = []

    for aba in abas:
        alteracoes = 0

        for linha in aba.iter_rows():
            for celula in linha:
                valor = celula.value

                if not parece_preco_excel(valor):
                    continue

                valor_decimal = converter_valor_excel_para_decimal(
                    valor
                )

                if valor_decimal is None:
                    continue

                novo_valor = (
                    valor_decimal * Decimal(str(multiplicador))
                )

                celula.value = arredondar_sempre_para_cima(
                    novo_valor
                )
                celula.number_format = '"R$" #,##0'
                alteracoes += 1

        abas_processadas.append({
            "Aba": aba.title,
            "Células atualizadas": alteracoes,
        })

    resultado = io.BytesIO()
    arquivo.save(resultado)

    return resultado.getvalue(), abas_processadas


# ==========================================================
# PDF -> EXCEL
# ==========================================================

def extrair_tabelas_pdf(arquivo_pdf):
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    registros = []

    for numero_pagina, pagina in enumerate(documento, start=1):
        texto = pagina.get_text("text")

        for linha in texto.splitlines():
            linha = linha.strip()

            if linha:
                registros.append({
                    "Página": numero_pagina,
                    "Conteúdo": linha,
                })

    documento.close()

    return pd.DataFrame(
        registros,
        columns=["Página", "Conteúdo"],
    )


def converter_pdf_para_excel(arquivo_pdf):
    df = extrair_tabelas_pdf(arquivo_pdf)
    saida = io.BytesIO()

    with pd.ExcelWriter(
        saida,
        engine="openpyxl",
    ) as escritor:
        df.to_excel(
            escritor,
            index=False,
            sheet_name="Dados extraídos",
        )

        planilha = escritor.sheets["Dados extraídos"]
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions
        planilha.column_dimensions["A"].width = 12
        planilha.column_dimensions["B"].width = 100

        for celulas in planilha.iter_rows():
            for celula in celulas:
                celula.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

                if celula.row == 1:
                    celula.font = Font(
                        bold=True,
                        color="FFFFFF",
                    )
                    celula.fill = PatternFill(
                        "solid",
                        fgColor="243B53",
                    )

    return saida.getvalue(), len(df)

# ==========================================================
# FUNÇÕES DE APOIO
# ==========================================================

def nome_arquivo_sem_extensao(nome):
    return Path(nome).stem


def criar_nome_saida(nome_original, sufixo, extensao):
    return (
        f"{nome_arquivo_sem_extensao(nome_original)}"
        f"_{sufixo}.{extensao}"
    )


def mostrar_resumo_pdf(conferencias):
    if not conferencias:
        st.warning(
            "Nenhum preço foi identificado na coluna TAB 1. "
            "Confira se o PDF contém texto selecionável."
        )
        return

    df = pd.DataFrame(conferencias)

    st.success(
        f"Processamento concluído: {len(df)} preço(s) atualizado(s)."
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )


def mostrar_resumo_excel(abas_processadas):
    if not abas_processadas:
        st.info("Nenhuma aba foi processada.")
        return

    df = pd.DataFrame(abas_processadas)
    total = int(df["Células atualizadas"].sum())

    st.success(
        f"Processamento concluído: {total} célula(s) atualizada(s)."
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )


def validar_multiplicador(multiplicador):
    try:
        numero = Decimal(str(multiplicador))
    except (InvalidOperation, ValueError):
        return False

    return numero.is_finite() and numero > 0


def validar_pdf(arquivo):
    return (
        arquivo is not None
        and arquivo.name.lower().endswith(".pdf")
    )


def validar_excel(arquivo):
    if arquivo is None:
        return False

    nome = arquivo.name.lower()
    return nome.endswith((".xlsx", ".xlsm"))


def validar_arquivo_nao_vazio(arquivo):
    return (
        arquivo is not None
        and len(arquivo.getvalue()) > 0
    )


def disponibilizar_download(
    conteudo,
    nome_arquivo,
    tipo_mime,
    texto_botao="Baixar arquivo processado",
):
    st.download_button(
        label=texto_botao,
        data=conteudo,
        file_name=nome_arquivo,
        mime=tipo_mime,
        use_container_width=True,
    )


# ==========================================================
# OPÇÕES DO PDF
# ==========================================================

def exibir_opcoes_pdf():
    st.subheader("Configurações do PDF")

    multiplicador = st.number_input(
        "Multiplicador dos preços da TAB 1",
        min_value=0.000001,
        value=2.383949988,
        step=0.1,
        format="%.9f",
        help=(
            "Cada preço da TAB 1 será multiplicado "
            "por este valor e arredondado para cima."
        ),
        key="multiplicador_pdf",
    )

    st.caption(
        "As demais colunas não terão seus preços atualizados."
    )

    remover_tab2 = st.checkbox(
        "Remover TAB 2",
        value=True,
        key="remover_tab2",
    )

    remover_tab3 = st.checkbox(
        "Remover TAB 3",
        value=True,
        key="remover_tab3",
    )

    formato = st.radio(
        "Formato do catálogo para impressão",
        [
            "Manter layout original",
            "A4 horizontal + letras maiores",
        ],
        index=0,
        key="formato_catalogo",
    )

    colunas_remover = []

    if remover_tab2:
        colunas_remover.append(2)

    if remover_tab3:
        colunas_remover.append(3)

    modo_horizontal = (
        formato == "A4 horizontal + letras maiores"
    )

    return (
        multiplicador,
        tuple(colunas_remover),
        modo_horizontal,
    )


# ==========================================================
# PDF -> PDF
# ==========================================================

def interface_pdf_para_pdf():
    st.header("📄 Atualizar preços em PDF")

    st.write(
        "Envie seu catálogo, escolha o multiplicador "
        "e gere uma nova versão com os preços atualizados."
    )

    arquivo_pdf = st.file_uploader(
        "Selecione o catálogo em PDF",
        type=["pdf"],
        key="upload_pdf_atualizar",
    )

    (
        multiplicador,
        colunas_remover,
        modo_horizontal,
    ) = exibir_opcoes_pdf()

    if arquivo_pdf is None:
        st.info("Envie um arquivo PDF para começar.")
        return

    if not validar_arquivo_nao_vazio(arquivo_pdf):
        st.error("O arquivo enviado está vazio.")
        return

    if not validar_pdf(arquivo_pdf):
        st.error("Selecione um arquivo PDF válido.")
        return

    if not validar_multiplicador(multiplicador):
        st.error("Informe um multiplicador maior que zero.")
        return

    if st.button(
        "🚀 Atualizar preços do PDF",
        type="primary",
        use_container_width=True,
        key="botao_atualizar_pdf",
    ):
        try:
            with st.spinner("Processando o catálogo..."):
                conteudo_pdf, conferencias = (
                    atualizar_pdf_mantendo_layout(
                        arquivo_pdf.getvalue(),
                        multiplicador,
                        deixar_fundo_branco=True,
                        colunas_remover=colunas_remover,
                        modo_horizontal=modo_horizontal,
                    )
                )

            st.session_state["pdf_processado"] = conteudo_pdf
            st.session_state["pdf_conferencias"] = conferencias
            st.session_state["pdf_nome_saida"] = criar_nome_saida(
                arquivo_pdf.name,
                "atualizado",
                "pdf",
            )

        except Exception as erro:
            st.error(f"Não foi possível processar o PDF: {erro}")
            return

    if "pdf_processado" in st.session_state:
        disponibilizar_download(
            st.session_state["pdf_processado"],
            st.session_state.get(
                "pdf_nome_saida",
                "catalogo_atualizado.pdf",
            ),
            "application/pdf",
            "⬇️ Baixar PDF atualizado",
        )

        mostrar_resumo_pdf(
            st.session_state.get("pdf_conferencias", [])
        )
        # ==========================================================
# PDF -> EXCEL
# ==========================================================

def interface_pdf_para_excel():
    st.header("📊 Converter PDF para Excel")

    st.write(
        "Extraia o texto do catálogo PDF e exporte "
        "o conteúdo para uma planilha Excel."
    )

    arquivo_pdf = st.file_uploader(
        "Selecione o arquivo PDF",
        type=["pdf"],
        key="upload_pdf_excel",
    )

    if arquivo_pdf is None:
        st.info("Envie um arquivo PDF para começar.")
        return

    if not validar_arquivo_nao_vazio(arquivo_pdf):
        st.error("O arquivo enviado está vazio.")
        return

    if not validar_pdf(arquivo_pdf):
        st.error("Selecione um arquivo PDF válido.")
        return

    if st.button(
        "📥 Converter para Excel",
        type="primary",
        use_container_width=True,
        key="botao_pdf_excel",
    ):
        try:
            with st.spinner("Extraindo o conteúdo do PDF..."):
                conteudo_excel, quantidade_linhas = (
                    converter_pdf_para_excel(
                        arquivo_pdf.getvalue()
                    )
                )

            st.session_state["excel_convertido"] = conteudo_excel
            st.session_state["excel_convertido_nome"] = (
                criar_nome_saida(
                    arquivo_pdf.name,
                    "convertido",
                    "xlsx",
                )
            )
            st.session_state["excel_convertido_linhas"] = (
                quantidade_linhas
            )

        except Exception as erro:
            st.error(f"Não foi possível converter o PDF: {erro}")
            return

    if "excel_convertido" in st.session_state:
        st.success(
            "Conversão concluída: "
            f"{st.session_state.get('excel_convertido_linhas', 0)} "
            "linhas extraídas."
        )

        disponibilizar_download(
            st.session_state["excel_convertido"],
            st.session_state.get(
                "excel_convertido_nome",
                "dados_convertidos.xlsx",
            ),
            (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "⬇️ Baixar Excel convertido",
        )


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def interface_excel_para_excel():
    st.header("📗 Atualizar preços em Excel")

    st.write(
        "Envie uma planilha Excel e multiplique os valores "
        "numéricos encontrados nas células."
    )

    arquivo_excel = st.file_uploader(
        "Selecione a planilha Excel",
        type=["xlsx", "xlsm"],
        key="upload_excel_atualizar",
    )

    multiplicador = st.number_input(
        "Multiplicador dos preços",
        min_value=0.000001,
        value=2.0,
        step=0.1,
        format="%.6f",
        key="multiplicador_excel",
    )

    if arquivo_excel is None:
        st.info("Envie uma planilha Excel para começar.")
        return

    if not validar_arquivo_nao_vazio(arquivo_excel):
        st.error("O arquivo enviado está vazio.")
        return

    if not validar_excel(arquivo_excel):
        st.error("Selecione um arquivo Excel válido.")
        return

    if not validar_multiplicador(multiplicador):
        st.error("Informe um multiplicador maior que zero.")
        return

    if st.button(
        "🚀 Atualizar planilha",
        type="primary",
        use_container_width=True,
        key="botao_atualizar_excel",
    ):
        try:
            with st.spinner("Atualizando a planilha..."):
                conteudo_excel, resumo = (
                    atualizar_planilha_excel(
                        arquivo_excel.getvalue(),
                        multiplicador,
                    )
                )

            st.session_state["planilha_atualizada"] = (
                conteudo_excel
            )
            st.session_state["planilha_resumo"] = resumo
            st.session_state["planilha_nome_saida"] = (
                criar_nome_saida(
                    arquivo_excel.name,
                    "atualizado",
                    "xlsx",
                )
            )

        except Exception as erro:
            st.error(
                f"Não foi possível atualizar o Excel: {erro}"
            )
            return

    if "planilha_atualizada" in st.session_state:
        disponibilizar_download(
            st.session_state["planilha_atualizada"],
            st.session_state.get(
                "planilha_nome_saida",
                "planilha_atualizada.xlsx",
            ),
            (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "⬇️ Baixar planilha atualizada",
        )

        mostrar_resumo_excel(
            st.session_state.get("planilha_resumo", [])
        )


# ==========================================================
# MENU PRINCIPAL
# ==========================================================

def main():
    st.sidebar.title("📋 Menu")

    opcao = st.sidebar.radio(
        "Escolha uma funcionalidade:",
        [
            "PDF → PDF",
            "PDF → Excel",
            "Excel → Excel",
        ],
        key="menu_principal",
    )

    st.sidebar.divider()

    st.sidebar.caption("Atualizador de Preços")

    st.sidebar.info(
        "Envie seu arquivo, configure as opções "
        "e baixe o resultado processado."
    )

    if opcao == "PDF → PDF":
        interface_pdf_para_pdf()

    elif opcao == "PDF → Excel":
        interface_pdf_para_excel()

    elif opcao == "Excel → Excel":
        interface_excel_para_excel()


# ==========================================================
# EXECUÇÃO
# ==========================================================

if __name__ == "__main__":
    main()
