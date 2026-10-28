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
st.caption(
    "PDF → PDF, PDF → Excel e Excel → Excel."
)


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
    """
    Converte:
    2210 -> R$ 2.210
    """
    numero = f"{valor_inteiro:,}"
    numero = numero.replace(",", ".")
    return f"R$ {numero}"


def calcular_novo_preco(texto_preco, multiplicador):
    """
    Novo preço = teto(preço original x multiplicador)
    """
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
    """
    Remove cabeçalhos e preços das colunas escolhidas.
    Mantém a largura da página para não cortar descrições.
    """
    colunas_remover = set(colunas_remover or [])

    if not colunas_remover:
        return

    centros = localizar_cabecalhos_tabelas(pagina)

    if not centros:
        return

    caixas_cabecalho = []
    estrutura = pagina.get_text("dict")

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
                    caixas_cabecalho.append(
                        pymupdf.Rect(span["bbox"])
                    )
    for caixa in caixas_cabecalho:
        pagina.add_redact_annot(
            pymupdf.Rect(
                caixa.x0 - 1.5,
                caixa.y0 - 0.5,
                caixa.x1 + 1.5,
                caixa.y1 + 0.5,
            ),
            fill=(0.75, 0.75, 0.75),
        )

    precos_remover = localizar_precos_na_pagina(
        pagina,
        tuple(colunas_remover),
    )

    for preco in precos_remover:
        caixa = preco["rect"]

        pagina.add_redact_annot(
            pymupdf.Rect(
                caixa.x0 - 1.5,
                caixa.y0 - 0.5,
                caixa.x1 + 1.5,
                caixa.y1 + 0.5,
            ),
            fill=(1, 1, 1),
        )

    if caixas_cabecalho or precos_remover:
        pagina.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )


def colocar_fundo_branco(pagina):
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
    colunas_remover=(),
):
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    conferencias = []

    if deixar_fundo_branco:
        for pagina in documento:
            colocar_fundo_branco(pagina)

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

            texto_novo = formatar_preco_brasileiro(
                novo_valor
            )

            caixa_original = preco["rect"]

            caixa_redacao = pymupdf.Rect(
                caixa_original.x0 - 1.5,
                caixa_original.y0 - 1.0,
                caixa_original.x1 + 1.5,
                caixa_original.y1 + 1.0,
            )

            pagina.add_redact_annot(
                caixa_redacao,
                fill=(1, 1, 1),
            )

            substituicoes.append({
                "texto_original": texto_original,
                "texto_novo": texto_novo,
                "caixa_original": caixa_original,
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
            caixa = item["caixa_original"]
            texto_novo = item["texto_novo"]
            fonte = item["fonte"]
            cor = item["cor"]
            tamanho_ajustado = item["tamanho"]

            largura_original = caixa.width

            largura_nova = medir_largura_texto(
                texto_novo,
                fonte,
                tamanho_ajustado,
            )

            while (
                largura_nova > largura_original
                and tamanho_ajustado > 5
            ):
                tamanho_ajustado -= 0.2

                largura_nova = medir_largura_texto(
                    texto_novo,
                    fonte,
                    tamanho_ajustado,
                )

            caixa_texto = pymupdf.Rect(
                caixa.x0,
                caixa.y0 - 0.5,
                max(
                    caixa.x1,
                    caixa.x0 + largura_nova + 3,
                ),
                caixa.y1 + 2,
            )

            resultado_texto = pagina.insert_textbox(
                caixa_texto,
                texto_novo,
                fontname=fonte,
                fontsize=tamanho_ajustado,
                color=cor,
                align=pymupdf.TEXT_ALIGN_LEFT,
                overlay=True,
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

            conferencias.append({
                "Página": numero_pagina,
                "Preço original": item["texto_original"],
                "Preço atualizado": item["texto_novo"],
            })

        remover_colunas_pdf(
            pagina,
            colunas_remover,
        )

    # Salva primeiro as alterações de preços e das colunas.
    pdf_intermediario = io.BytesIO()
    documento.save(
        pdf_intermediario,
        garbage=4,
        deflate=True,
    )
    documento.close()

    # Amplia o catálogo em 20% e distribui cada página original em até
    # quatro folhas do mesmo tamanho. Assim, a impressora não reduz o PDF
    # para caber na página e o conteúdo não é descartado: as partes que
    # ultrapassam a primeira folha continuam nas folhas seguintes.
    fator_ampliacao = 1.20
    pdf_intermediario.seek(0)
    documento_original = pymupdf.open(
        stream=pdf_intermediario.getvalue(),
        filetype="pdf",
    )
    documento_ampliado = pymupdf.open()

    for pagina_original in documento_original:
        largura_original = pagina_original.rect.width
        altura_original = pagina_original.rect.height
        largura_ampliada = largura_original * fator_ampliacao
        altura_ampliada = altura_original * fator_ampliacao

        # Cria uma página virtual ampliada mantendo as proporções.
        pagina_virtual = pymupdf.open()
        pagina_grande = pagina_virtual.new_page(
            width=largura_ampliada,
            height=altura_ampliada,
        )
        pagina_grande.show_pdf_page(
            pagina_grande.rect,
            documento_original,
            pagina_original.number,
            keep_proportion=True,
        )

        # Divide a página ampliada em folhas do tamanho original (A4).
        # As folhas de borda só são criadas se houver conteúdo além delas.
        colunas = int((largura_ampliada + largura_original - 1) // largura_original)
        linhas = int((altura_ampliada + altura_original - 1) // altura_original)

        for linha in range(linhas):
            for coluna in range(colunas):
                x0 = coluna * largura_original
                y0 = linha * altura_original
                x1 = min(x0 + largura_original, largura_ampliada)
                y1 = min(y0 + altura_original, altura_ampliada)
                pagina_saida = documento_ampliado.new_page(
                    width=largura_original,
                    height=altura_original,
                )
                # Mantém a escala de 120%; a folha recorta naturalmente
                # apenas o trecho que fica fora de sua área imprimível.
                destino = pymupdf.Rect(
                    -x0,
                    -y0,
                    largura_ampliada - x0,
                    altura_ampliada - y0,
                )
                pagina_saida.show_pdf_page(
                    destino,
                    pagina_virtual,
                    0,
                    keep_proportion=False,
                )

        pagina_virtual.close()

    arquivo_final = io.BytesIO()
    documento_ampliado.save(
        arquivo_final,
        garbage=4,
        deflate=True,
    )
    documento_ampliado.close()
    documento_original.close()
    arquivo_final.seek(0)

    return arquivo_final.getvalue(), conferencias


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def converter_valor_excel_para_decimal(valor):
    """
    Aceita:
    926.82
    926,82
    1.073,88
    R$ 1.073,88
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
    Aceita números e preços.
    Evita alterar medidas e códigos.
    """
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
    arquivo = load_workbook(
        io.BytesIO(arquivo_excel),
    )

    abas_processadas = []

    if nome_aba and nome_aba in arquivo.sheetnames:
        abas = [arquivo[nome_aba]]
    else:
        abas = arquivo.worksheets

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
    resultado.seek(0)

    return resultado.getvalue(), abas_processadas

# ==========================================================
# PDF -> EXCEL
# ==========================================================

def extrair_tabelas_pdf(arquivo_pdf):
    """
    Extrai as tabelas de texto do PDF para um DataFrame.
    Mantém o texto encontrado em cada página.
    """
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
    """Converte o texto extraído do PDF para um arquivo Excel."""
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

    saida.seek(0)

    return saida.getvalue(), len(df)


# ==========================================================
# FUNÇÕES DE APOIO PARA A INTERFACE
# ==========================================================

def nome_arquivo_sem_extensao(nome):
    return Path(nome).stem


def criar_nome_saida(nome_original, sufixo, extensao):
    nome = nome_arquivo_sem_extensao(nome_original)

    return f"{nome}_{sufixo}.{extensao}"


def mostrar_resumo_pdf(conferencias):
    if not conferencias:
        st.warning(
            "Nenhum preço foi identificado na coluna TAB 1. "
            "Confira se o PDF contém texto selecionável."
        )
        return

    df = pd.DataFrame(conferencias)

    st.success(
        f"Processamento concluído: "
        f"{len(df)} preço(s) atualizado(s)."
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
        f"Processamento concluído: "
        f"{total} célula(s) atualizada(s)."
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )


# ==========================================================
# VALIDAÇÃO DOS ARQUIVOS
# ==========================================================

def validar_multiplicador(multiplicador):
    try:
        numero = Decimal(str(multiplicador))
    except (InvalidOperation, ValueError):
        return False

    return numero.is_finite() and numero > 0


def validar_pdf(arquivo):
    if arquivo is None:
        return False

    return arquivo.name.lower().endswith(".pdf")


def validar_excel(arquivo):
    if arquivo is None:
        return False

    nome = arquivo.name.lower()

    return nome.endswith(".xlsx") or nome.endswith(".xlsm")


def validar_arquivo_nao_vazio(arquivo):
    if arquivo is None:
        return False

    conteudo = arquivo.getvalue()

    return len(conteudo) > 0


# ==========================================================
# DOWNLOADS
# ==========================================================

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
# CONFIGURAÇÕES DO PROCESSAMENTO DE PDF
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

    colunas_remover = []

    if remover_tab2:
        colunas_remover.append(2)

    if remover_tab3:
        colunas_remover.append(3)

    return multiplicador, tuple(colunas_remover)


# ==========================================================
# CONVERSÃO DE DADOS PARA DOWNLOAD
# ==========================================================

def gerar_excel_com_dataframe(df, nome_aba="Dados"):
    saida = io.BytesIO()

    with pd.ExcelWriter(
        saida,
        engine="openpyxl",
    ) as escritor:
        df.to_excel(
            escritor,
            index=False,
            sheet_name=nome_aba[:31],
        )

        planilha = escritor.sheets[nome_aba[:31]]
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        for coluna in planilha.columns:
            letra = get_column_letter(coluna[0].column)

            maior = max(
                (
                    len(str(celula.value))
                    for celula in coluna
                    if celula.value is not None
                ),
                default=10,
            )

            planilha.column_dimensions[letra].width = min(
                maior + 2,
                60,
            )

        for celulas in planilha.iter_rows():
            for celula in celulas:
                celula.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

                if celula.row == 1:
                    celula.font = Font(bold=True)

    saida.seek(0)

    return saida.getvalue()

# ==========================================================
# INTERFACE PRINCIPAL
# ==========================================================

def interface_pdf_para_pdf():
    st.header("📄 Atualizar preços em PDF")

    st.write(
        "Envie seu catálogo PDF, escolha o multiplicador "
        "e gere uma nova versão com os preços atualizados."
    )

    arquivo_pdf = st.file_uploader(
        "Selecione o catálogo em PDF",
        type=["pdf"],
        key="upload_pdf_atualizar",
    )

    multiplicador, colunas_remover = exibir_opcoes_pdf()

    if arquivo_pdf is None:
        st.info("Envie um arquivo PDF para começar.")
        return

    if not validar_arquivo_nao_vazio(arquivo_pdf):
        st.error("O arquivo enviado está vazio.")
        return

    if not validar_pdf(arquivo_pdf):
        st.error("Selecione um arquivo PDF válido.")
        return

    st.caption(f"Arquivo selecionado: {arquivo_pdf.name}")

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
            st.error(
                f"Não foi possível processar o PDF: {erro}"
            )
            return

    if "pdf_processado" in st.session_state:
        conteudo_pdf = st.session_state["pdf_processado"]

        disponibilizar_download(
            conteudo_pdf,
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
            st.error(
                f"Não foi possível converter o PDF: {erro}"
            )
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
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet",
            "⬇️ Baixar Excel convertido",
        )


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
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet",
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

    st.sidebar.caption(
        "Atualizador de Preços"
    )

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
    
