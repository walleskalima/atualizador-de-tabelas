import io
import re
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

CINZA_CABECALHO = (0.75, 0.75, 0.75)


# ==========================================================
# PREÇOS E ARREDONDAMENTO
# ==========================================================

# Aceita "926,82", "1.073,88", "1073,88" e "R$ 926,82".
# Os lookarounds impedem casar só um pedaço de um número maior.
PADRAO_PRECO = re.compile(
    r"(?<![\d.,])(?:R\$\s*)?\d+(?:\.\d{3})*,\d{2}(?!\d)"
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
    2210 -> R$ 2.210,00
    """
    numero = f"{valor_inteiro:,}"
    numero = numero.replace(",", ".")
    return f"R$ {numero},00"


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


def escolher_fonte_pdf(nome_fonte_original, flags=0):
    """Escolhe a fonte base-14 equivalente, detectando negrito pelo
    nome da fonte OU pelas flags do span (bit 16 = bold)."""
    nome = str(nome_fonte_original).lower()
    negrito = (
        "bold" in nome
        or "black" in nome
        or "heavy" in nome
        or bool(int(flags or 0) & 16)
    )

    if "times" in nome:
        return "tibo" if negrito else "tiro"

    if "courier" in nome:
        return "cobo" if negrito else "cour"

    return "hebo" if negrito else "helv"


def medir_largura_texto(texto, fonte, tamanho):
    try:
        return pymupdf.get_text_length(
            texto,
            fontname=fonte,
            fontsize=tamanho,
        )
    except Exception:
        return len(texto) * tamanho * 0.6


def inserir_preco(
    pagina,
    texto,
    fonte,
    tamanho,
    cor,
    baseline_y,
    x_min,
    x_max,
    alinhar="dir",
):
    """Escreve o preço na linha de base original.

    Mede o texto e, se não couber entre x_min e x_max, diminui a fonte
    em vez de deixar o número ser cortado.
    alinhar="dir": a borda direita do texto fica em x_max.
    alinhar="esq": a borda esquerda do texto fica em x_min.
    """
    largura_disponivel = max(x_max - x_min, 1)
    largura = medir_largura_texto(texto, fonte, tamanho)

    while largura > largura_disponivel and tamanho > 6:
        tamanho -= 0.5
        largura = medir_largura_texto(texto, fonte, tamanho)

    if alinhar == "dir":
        x = x_max - largura
    else:
        x = x_min

    pagina.insert_text(
        pymupdf.Point(x, baseline_y),
        texto,
        fontname=fonte,
        fontsize=tamanho,
        color=cor,
        overlay=True,
    )


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


def calcular_limites_colunas(pagina, centros):
    """Limites horizontais de cada coluna TAB, a partir dos cabeçalhos."""
    limites = {}
    centros_ordenados = sorted(centros.items(), key=lambda item: item[1])

    for indice, (numero, centro) in enumerate(centros_ordenados):
        if indice > 0:
            esquerda = (centros_ordenados[indice - 1][1] + centro) / 2
        elif len(centros_ordenados) > 1:
            distancia_direita = centros_ordenados[1][1] - centro
            esquerda = centro - distancia_direita / 2
        else:
            esquerda = centro - 25

        if indice < len(centros_ordenados) - 1:
            direita = (centro + centros_ordenados[indice + 1][1]) / 2
        elif len(centros_ordenados) > 1:
            distancia_esquerda = centro - centros_ordenados[indice - 1][1]
            direita = centro + distancia_esquerda / 2
        else:
            direita = centro + 25

        limites[numero] = (max(0, esquerda), min(pagina.rect.width, direita))

    return limites


def localizar_precos_na_pagina(pagina, colunas_precos=(1,), centros=None):
    """Localiza preços pelas coordenadas das colunas TAB do PDF.

    Usa 'rawdict' para obter a caixa EXATA do preço (por caractere),
    em vez de estimar por proporção de letras.
    """
    encontrados = []

    if centros is None:
        centros = localizar_cabecalhos_tabelas(pagina)

    if not centros:
        return encontrados

    limites = calcular_limites_colunas(pagina, centros)
    estrutura = pagina.get_text("rawdict")

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue
        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                chars = span.get("chars", [])
                if not chars:
                    continue

                texto_span = "".join(c["c"] for c in chars)

                for correspondencia in PADRAO_PRECO.finditer(texto_span):
                    inicio, fim = correspondencia.span()

                    x0 = chars[inicio]["bbox"][0]
                    x1 = chars[fim - 1]["bbox"][2]
                    y0 = span["bbox"][1]
                    y1 = span["bbox"][3]
                    caixa_preco = pymupdf.Rect(x0, y0, x1, y1)

                    centro_preco = (x0 + x1) / 2
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
                        "flags": span.get("flags", 0),
                        "tamanho_fonte": span.get("size", 10),
                        "cor": rgb_inteiro_para_tupla(span.get("color", 0)),
                        "baseline_y": span["origin"][1],
                        "x_esquerda_coluna": limites[coluna_encontrada][0],
                        "x_direita_coluna": limites[coluna_encontrada][1],
                    })

    return encontrados


def limite_direito_conteudo(pagina, margem=12):
    """Maior x ocupado por texto na página + margem."""
    xs = [palavra[2] for palavra in pagina.get_text("words")]
    if not xs:
        return pagina.rect.width
    return min(pagina.rect.width, max(xs) + margem)


def remover_colunas_pdf(pagina, colunas_remover):
    """Apaga cabeçalhos/preços das colunas escolhidas e recorta a página
    no limite real do conteúdo quando as colunas removidas são as finais."""
    colunas_remover = set(colunas_remover or [])
    if not colunas_remover:
        return

    centros = localizar_cabecalhos_tabelas(pagina)
    if not centros:
        return

    estrutura = pagina.get_text("dict")
    caixas_cabecalho = []
    caixas_preco = []
    cabecalhos_tab1 = []

    for bloco in estrutura.get("blocks", []):
        if bloco.get("type") != 0:
            continue
        for linha in bloco.get("lines", []):
            for span in linha.get("spans", []):
                texto = span.get("text", "").strip()
                m = re.fullmatch(r"TAB\s*([123])", texto, re.IGNORECASE)
                if not m:
                    continue
                numero = int(m.group(1))
                if numero in colunas_remover:
                    caixas_cabecalho.append(pymupdf.Rect(span["bbox"]))
                if numero == 1:
                    cabecalhos_tab1.append(pymupdf.Rect(span["bbox"]))

    for preco in localizar_precos_na_pagina(
        pagina, tuple(colunas_remover), centros=centros
    ):
        caixas_preco.append(preco["rect"])

    # Cabeçalhos são apagados com o MESMO cinza da faixa, para não
    # deixar um retângulo branco; preços são apagados em branco.
    for caixa in caixas_cabecalho:
        pagina.add_redact_annot(
            pymupdf.Rect(caixa.x0 - 1, caixa.y0 - 1,
                         caixa.x1 + 1, caixa.y1 + 1),
            fill=CINZA_CABECALHO,
        )

    for caixa in caixas_preco:
        pagina.add_redact_annot(
            pymupdf.Rect(caixa.x0 - 1, caixa.y0 - 0.5,
                         caixa.x1 + 1, caixa.y1 + 0.5),
            fill=(1, 1, 1),
        )

    if caixas_cabecalho or caixas_preco:
        pagina.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
        )

    colunas_mantidas = [n for n in centros if n not in colunas_remover]
    if not colunas_mantidas:
        return

    ultima_mantida = max(colunas_mantidas, key=lambda n: centros[n])
    so_removeu_a_direita = all(
        n in colunas_remover
        for n in centros
        if centros[n] > centros[ultima_mantida]
    )

    if not so_removeu_a_direita:
        return

    # Limite pelo conteúdo real (preços novos já inseridos), nunca "centro + N".
    limite_direito = limite_direito_conteudo(pagina)
    if limite_direito <= 100:
        return

    pagina.set_cropbox(
        pymupdf.Rect(0, 0, limite_direito, pagina.rect.height)
    )

    # Garante que a faixa cinza do cabeçalho vai até a borda final.
    for cabecalho in cabecalhos_tab1:
        if cabecalho.x1 + 2 < limite_direito:
            pagina.draw_rect(
                pymupdf.Rect(
                    cabecalho.x1 + 2,
                    cabecalho.y0 - 1,
                    limite_direito,
                    cabecalho.y1 + 1,
                ),
                color=None,
                fill=CINZA_CABECALHO,
                overlay=False,
            )


def colocar_fundo_branco(pagina):
    pagina.draw_rect(
        pagina.rect,
        color=None,
        fill=(1, 1, 1),
        overlay=False,
    )


def achar_corte(pagina_origem, altura, largura):
    """Escolhe a altura de corte da página.

    1) Prefere uma linha horizontal separadora entre produtos.
    2) Senão, usa o maior espaço vazio (sem texto e sem imagem) perto do meio.
    3) Só em último caso corta em altura/2.
    """
    separadores = []

    for desenho in pagina_origem.get_drawings():
        for item in desenho.get("items", []):
            if item[0] == "l":
                a, b = item[1], item[2]
                if abs(a.y - b.y) <= 1.5 and abs(a.x - b.x) >= largura * 0.72:
                    y = (a.y + b.y) / 2
                    if altura * 0.30 < y < altura * 0.72:
                        separadores.append(y)

    if separadores:
        return min(separadores, key=lambda y: abs(y - altura / 2))

    faixas = [(w[1], w[3]) for w in pagina_origem.get_text("words")]
    try:
        for info in pagina_origem.get_image_info():
            bbox = info.get("bbox")
            if bbox:
                faixas.append((bbox[1], bbox[3]))
    except Exception:
        pass

    faixas.sort()
    lacunas = []
    fim = 0.0
    for y0, y1 in faixas:
        if y0 > fim + 2:
            lacunas.append((fim + y0) / 2)
        fim = max(fim, y1)

    candidatos = [y for y in lacunas if altura * 0.30 < y < altura * 0.70]
    if candidatos:
        return min(candidatos, key=lambda y: abs(y - altura / 2))

    return altura / 2


def ampliar_catalogo_para_impressao(pdf_bytes):
    """Reorganiza cada página original em duas páginas A4 horizontais.

    O corte é feito sobre um separador/espaço vazio entre produtos, o
    conteúdo é ampliado para ocupar a folha e a fonte fica maior.
    """
    origem = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    saida = pymupdf.open()
    largura_pagina = 842  # A4 horizontal, pontos PDF
    altura_pagina = 595
    margem = 18

    for pagina_origem in origem:
        largura = pagina_origem.rect.width
        altura = pagina_origem.rect.height
        corte = achar_corte(pagina_origem, altura, largura)

        faixas = [(0.0, corte), (corte, altura)]
        for y0, y1 in faixas:
            clip = pymupdf.Rect(0, max(0, y0), largura, min(altura, y1))
            if clip.height < 80:
                continue
            texto_faixa = pagina_origem.get_text("text", clip=clip).strip()
            if not texto_faixa:
                continue

            pagina_nova = saida.new_page(width=largura_pagina, height=altura_pagina)
            escala = min((largura_pagina - 2 * margem) / clip.width,
                         (altura_pagina - 2 * margem) / clip.height)
            w = clip.width * escala
            h = clip.height * escala
            x0_destino = (largura_pagina - w) / 2
            y0_destino = (altura_pagina - h) / 2
            destino = pymupdf.Rect(x0_destino, y0_destino,
                                   x0_destino + w, y0_destino + h)
            pagina_nova.show_pdf_page(destino, origem, pagina_origem.number,
                                      clip=clip, keep_proportion=True)

    resultado = saida.tobytes(garbage=4, deflate=True)
    saida.close()
    origem.close()
    return resultado


def atualizar_pdf_mantendo_layout(
    arquivo_pdf,
    multiplicador,
    deixar_fundo_branco=True,
    colunas_remover=(),
    formato_saida="Manter layout original",
):
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    conferencias = []
    aumentar_preco = formato_saida == "A4 horizontal + letras maiores"

    if deixar_fundo_branco:
        for pagina in documento:
            colocar_fundo_branco(pagina)

    for numero_pagina, pagina in enumerate(documento, start=1):
        # Cabeçalhos calculados UMA vez por página.
        centros = localizar_cabecalhos_tabelas(pagina)

        precos_pagina = localizar_precos_na_pagina(
            pagina, colunas_precos=(1,), centros=centros
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
            caixa_original = preco["rect"]

            pagina.add_redact_annot(
                pymupdf.Rect(
                    caixa_original.x0 - 1.5,
                    caixa_original.y0 - 1.0,
                    caixa_original.x1 + 1.5,
                    caixa_original.y1 + 1.0,
                ),
                fill=(1, 1, 1),
            )

            substituicoes.append(
                {
                    "texto_original": texto_original,
                    "texto_novo": texto_novo,
                    "caixa_original": caixa_original,
                    "fonte": escolher_fonte_pdf(
                        preco["fonte_original"], preco["flags"]
                    ),
                    "tamanho": preco["tamanho_fonte"],
                    "cor": preco["cor"],
                    "baseline_y": preco["baseline_y"],
                    "x_esquerda_coluna": preco["x_esquerda_coluna"],
                    "x_direita_coluna": preco["x_direita_coluna"],
                }
            )

        if substituicoes:
            pagina.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,
                graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                text=pymupdf.PDF_REDACT_TEXT_REMOVE,
            )

        for item in substituicoes:
            caixa = item["caixa_original"]

            # No modo horizontal o valor fica 1,5x maior (leitura facilitada).
            tamanho = item["tamanho"] * (1.5 if aumentar_preco else 1.0)

            if aumentar_preco:
                # Alinhado à direita, como no original; se não couber na
                # coluna, a fonte diminui em vez de cortar o último dígito.
                inserir_preco(
                    pagina,
                    item["texto_novo"],
                    item["fonte"],
                    tamanho,
                    item["cor"],
                    baseline_y=item["baseline_y"],
                    x_min=item["x_esquerda_coluna"],
                    x_max=min(caixa.x1 + 4, item["x_direita_coluna"] - 2),
                    alinhar="dir",
                )
            else:
                inserir_preco(
                    pagina,
                    item["texto_novo"],
                    item["fonte"],
                    tamanho,
                    item["cor"],
                    baseline_y=item["baseline_y"],
                    x_min=caixa.x0,
                    x_max=item["x_direita_coluna"] - 1,
                    alinhar="esq",
                )

            conferencias.append(
                {
                    "Página": numero_pagina,
                    "Preço original": item["texto_original"],
                    "Preço atualizado": item["texto_novo"],
                }
            )

        remover_colunas_pdf(pagina, colunas_remover)

    arquivo_final = io.BytesIO()

    documento.save(
        arquivo_final,
        garbage=4,
        deflate=True,
    )

    documento.close()
    arquivo_final.seek(0)

    pdf_processado = arquivo_final.getvalue()

    if aumentar_preco:
        pdf_processado = ampliar_catalogo_para_impressao(pdf_processado)

    return pdf_processado, conferencias


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

    if isinstance(valor, bool):
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
    if isinstance(valor, bool):
        return False

    if isinstance(valor, (int, float, Decimal)):
        return True

    if not isinstance(valor, str):
        return False

    padrao = (
        r"^\s*"
        r"(R\$\s*)?"
        r"\d{1,3}(?:\.\d{3})*"
        r"(?:,\d{1,2})?"
        r"\s*$"
    )

    return bool(re.match(padrao, valor))


def e_formula(valor):
    return isinstance(valor, str) and valor.startswith("=")


def manter_formula_original(formula_original):
    """
    Mantém a fórmula original da planilha sem multiplicá-la.

    O multiplicador é aplicado somente aos valores digitados
    diretamente nas células, e não às fórmulas.
    """
    return formula_original


def atualizar_excel(
    arquivo_excel,
    nome_aba,
    linha_inicio,
    numeros_colunas,
    multiplicador,
):
    """
    Atualiza várias colunas escolhidas.

    Valores diretos:
    teto(valor x multiplicador)

    Fórmulas:
    mantém a fórmula original sem multiplicá-la.
    """

    workbook = load_workbook(
        io.BytesIO(arquivo_excel),
        data_only=False,
        keep_links=True,
        rich_text=True,
    )

    planilha = workbook[nome_aba]
    conferencias = []

    for numero_coluna in numeros_colunas:

        for numero_linha in range(
            linha_inicio,
            planilha.max_row + 1,
        ):
            celula = planilha.cell(
                numero_linha,
                numero_coluna,
            )

            valor_original = celula.value

            if valor_original is None:
                continue

            # Fórmula: mantém exatamente como está.
            if e_formula(valor_original):
                celula.value = manter_formula_original(valor_original)

                conferencias.append(
                    {
                        "Linha": numero_linha,
                        "Célula": celula.coordinate,
                        "Valor original": valor_original,
                        "Valor atualizado": valor_original,
                        "Tipo": "Fórmula mantida (NÃO multiplicada)",
                    }
                )

                continue

            # Valor direto
            if not parece_preco_excel(valor_original):
                continue

            valor_decimal = converter_valor_excel_para_decimal(
                valor_original
            )

            if valor_decimal is None:
                continue

            novo_valor = arredondar_sempre_para_cima(
                valor_decimal
                * Decimal(str(multiplicador))
            )

            if (
                isinstance(valor_original, str)
                and "R$" in valor_original.upper()
            ):
                celula.value = formatar_preco_brasileiro(novo_valor)
            else:
                celula.value = novo_valor

                formato_anterior = str(celula.number_format)

                if "R$" in formato_anterior or "$" in formato_anterior:
                    celula.number_format = "R$ #,##0"

            conferencias.append(
                {
                    "Linha": numero_linha,
                    "Célula": celula.coordinate,
                    "Valor original": valor_original,
                    "Valor atualizado": celula.value,
                    "Tipo": "Valor multiplicado",
                }
            )

    # Faz o Excel recalcular as fórmulas ao abrir
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"

    arquivo_final = io.BytesIO()

    workbook.save(arquivo_final)

    arquivo_final.seek(0)

    return arquivo_final.getvalue(), conferencias


# ==========================================================
# PDF -> EXCEL
# ==========================================================

def criar_cabecalho(planilha, titulos):
    for numero_coluna, titulo in enumerate(titulos, start=1):
        celula = planilha.cell(
            1,
            numero_coluna,
            titulo,
        )

        celula.font = Font(
            bold=True,
            color="FFFFFF",
        )

        celula.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )

        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def inserir_pagina_pdf_como_imagem(planilha, pagina):
    pix = pagina.get_pixmap(
        matrix=pymupdf.Matrix(1.4, 1.4),
        alpha=False,
    )

    imagem_bytes = pix.tobytes("png")

    # Em memória: não deixa arquivos temporários no disco.
    imagem_excel = ExcelImage(io.BytesIO(imagem_bytes))

    largura_maxima = 1100

    if imagem_excel.width > largura_maxima:
        proporcao = largura_maxima / imagem_excel.width

        imagem_excel.width = int(imagem_excel.width * proporcao)
        imagem_excel.height = int(imagem_excel.height * proporcao)

    planilha.add_image(imagem_excel, "A1")
    planilha.column_dimensions["A"].width = 20


def adicionar_tabela_extraida(
    planilha,
    tabela,
    linha_inicial,
    numero_pagina,
):
    dados = tabela.extract()

    if not dados:
        return linha_inicial

    celula_titulo = planilha.cell(
        linha_inicial,
        1,
        "Tabela encontrada - Página " + str(numero_pagina),
    )

    celula_titulo.font = Font(
        bold=True,
        color="FFFFFF",
    )

    celula_titulo.fill = PatternFill(
        fill_type="solid",
        fgColor="4F81BD",
    )

    linha_inicial += 1

    for linha_pdf in dados:
        for numero_coluna, valor in enumerate(
            linha_pdf,
            start=1,
        ):
            celula = planilha.cell(
                linha_inicial,
                numero_coluna,
                valor,
            )

            celula.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

        linha_inicial += 1

    return linha_inicial + 2


def converter_pdf_para_excel(arquivo_pdf):
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    workbook = Workbook()

    planilha_tabelas = workbook.active
    planilha_tabelas.title = "Tabelas extraídas"

    planilha_texto = workbook.create_sheet("Texto por página")

    criar_cabecalho(
        planilha_texto,
        [
            "Página",
            "Texto extraído",
        ],
    )

    planilha_texto.column_dimensions["A"].width = 12
    planilha_texto.column_dimensions["B"].width = 100
    planilha_texto.freeze_panes = "A2"

    linha_texto = 2
    linha_tabelas = 1
    quantidade_tabelas = 0

    for numero_pagina, pagina in enumerate(
        documento,
        start=1,
    ):
        texto_pagina = pagina.get_text("text")

        for linha in texto_pagina.splitlines():
            linha = linha.strip()

            if not linha:
                continue

            planilha_texto.cell(
                linha_texto,
                1,
                numero_pagina,
            )

            celula_texto = planilha_texto.cell(
                linha_texto,
                2,
                linha,
            )

            celula_texto.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

            linha_texto += 1

        try:
            resultado_tabelas = pagina.find_tables()

            for tabela in resultado_tabelas.tables:
                linha_tabelas = adicionar_tabela_extraida(
                    planilha_tabelas,
                    tabela,
                    linha_tabelas,
                    numero_pagina,
                )

                quantidade_tabelas += 1

        except Exception:
            pass

        planilha_visual = workbook.create_sheet(
            "Visual página " + str(numero_pagina)
        )

        inserir_pagina_pdf_como_imagem(
            planilha_visual,
            pagina,
        )

    planilha_tabelas.column_dimensions["A"].width = 32

    for numero_coluna in range(2, 20):
        letra = get_column_letter(numero_coluna)
        planilha_tabelas.column_dimensions[letra].width = 22

    planilha_info = workbook.create_sheet("Informações")

    planilha_info["A1"] = "Conversão de PDF para Excel"
    planilha_info["A1"].font = Font(
        bold=True,
        size=14,
    )

    planilha_info["A3"] = "Tabelas identificadas"
    planilha_info["B3"] = quantidade_tabelas

    planilha_info["A5"] = "Observação"
    planilha_info["A5"].font = Font(bold=True)

    planilha_info["B5"] = (
        "A aba 'Tabelas extraídas' contém tabelas reconhecidas. "
        "A aba 'Texto por página' contém conteúdo editável. "
        "As abas 'Visual página N' preservam a aparência das páginas "
        "originais como imagem."
    )

    planilha_info["B5"].alignment = Alignment(
        wrap_text=True,
        vertical="top",
    )

    planilha_info.column_dimensions["A"].width = 28
    planilha_info.column_dimensions["B"].width = 100
    planilha_info.row_dimensions[5].height = 75

    arquivo_final = io.BytesIO()
    workbook.save(arquivo_final)
    arquivo_final.seek(0)

    documento.close()

    return arquivo_final.getvalue(), quantidade_tabelas


# ==========================================================
# INTERFACE PRINCIPAL
# ==========================================================

MIME_XLSX = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

arquivo = st.file_uploader(
    "Envie um arquivo PDF ou Excel",
    type=["pdf", "xlsx"],
)

multiplicador = st.number_input(
    "Multiplicador de preço",
    min_value=0.000001,
    value=2.383949988,
    step=0.000001,
    format="%.9f",
)

if arquivo is None:
    st.info("Envie um arquivo para começar.")
    st.stop()

nome_arquivo = arquivo.name
extensao = Path(nome_arquivo).suffix.lower()
conteudo_arquivo = arquivo.getvalue()


# ==========================================================
# TELA PDF
# ==========================================================

if extensao == ".pdf":
    st.subheader("PDF")

    modo_pdf = st.radio(
        "O que deseja gerar?",
        [
            "PDF com preços atualizados",
            "Excel convertido do PDF",
        ],
    )

    if modo_pdf == "PDF com preços atualizados":
        deixar_fundo_branco = st.checkbox(
            "Adicionar fundo branco atrás dos elementos",
            value=True,
        )

        st.markdown("**Formato do PDF final**")
        formato_saida = st.radio(
            "Escolha como deseja organizar o catálogo para impressão:",
            [
                "Manter layout original",
                "A4 horizontal + letras maiores",
            ],
            index=1,
            help=(
                "A opção horizontal divide o conteúdo em blocos maiores em folhas A4 horizontais, "
                "para aproveitar melhor a largura e facilitar a leitura. Ela pode gerar mais páginas."
            ),
            key="formato_saida_pdf",
        )
        if formato_saida == "A4 horizontal + letras maiores":
            st.caption(
                "O catálogo será reorganizado em A4 horizontal. O conteúdo é dividido em blocos "
                "para caberem vários produtos por página, em vez de um produto por folha."
            )

        st.markdown("**Colunas que deseja remover do PDF**")
        st.caption("A TAB 1 continuará sendo multiplicada. As colunas finais removidas também serão recortadas da página.")
        remover_tab2 = st.checkbox("Remover TAB 2", value=True, key="remover_tab2_pdf")
        remover_tab3 = st.checkbox("Remover TAB 3", value=True, key="remover_tab3_pdf")
        colunas_remover = []
        if remover_tab2:
            colunas_remover.append(2)
        if remover_tab3:
            colunas_remover.append(3)

        if st.button(
            "Gerar PDF com preços atualizados",
            type="primary",
        ):
            try:
                with st.spinner("Atualizando preços no PDF..."):
                    resultado, conferencias = atualizar_pdf_mantendo_layout(
                        arquivo_pdf=conteudo_arquivo,
                        multiplicador=multiplicador,
                        deixar_fundo_branco=deixar_fundo_branco,
                        colunas_remover=colunas_remover,
                        formato_saida=formato_saida,
                    )

                st.session_state["res_pdf"] = {
                    "arquivo": nome_arquivo,
                    "pdf": resultado,
                    "conferencias": conferencias,
                    "colunas": list(colunas_remover),
                    "nome_saida": Path(nome_arquivo).stem + "_precos_atualizados.pdf",
                }

            except Exception as erro:
                st.session_state.pop("res_pdf", None)
                st.error(
                    "Não foi possível processar o PDF. "
                    "Detalhe técnico: " + str(erro)
                )

        # O resultado fica guardado, então o download não faz ele sumir.
        res_pdf = st.session_state.get("res_pdf")

        if res_pdf and res_pdf["arquivo"] == nome_arquivo:
            if not res_pdf["conferencias"]:
                st.warning(
                    "Nenhum preço foi identificado. "
                    "O PDF pode ser escaneado, não ter cabeçalhos TAB 1/2/3 "
                    "ou usar outro padrão de preço."
                )
            else:
                colunas_txt = (
                    " Colunas removidas: "
                    + ", ".join("TAB " + str(c) for c in res_pdf["colunas"])
                    + "."
                    if res_pdf["colunas"]
                    else " Nenhuma coluna removida."
                )

                st.success(
                    "PDF pronto. "
                    + str(len(res_pdf["conferencias"]))
                    + " preço(s) da TAB 1 atualizado(s)."
                    + colunas_txt
                )

                st.dataframe(
                    pd.DataFrame(res_pdf["conferencias"]),
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    label="⬇️ Baixar PDF atualizado",
                    data=res_pdf["pdf"],
                    file_name=res_pdf["nome_saida"],
                    mime="application/pdf",
                )

    elif modo_pdf == "Excel convertido do PDF":
        if st.button(
            "Converter PDF para Excel",
            type="primary",
        ):
            try:
                with st.spinner("Extraindo textos, tabelas e páginas..."):
                    resultado, quantidade_tabelas = converter_pdf_para_excel(
                        conteudo_arquivo
                    )

                st.session_state["res_conv"] = {
                    "arquivo": nome_arquivo,
                    "xlsx": resultado,
                    "tabelas": quantidade_tabelas,
                    "nome_saida": Path(nome_arquivo).stem + "_convertido.xlsx",
                }

            except Exception as erro:
                st.session_state.pop("res_conv", None)
                st.error(
                    "Não foi possível converter o PDF para Excel. "
                    "Detalhe técnico: " + str(erro)
                )

        res_conv = st.session_state.get("res_conv")

        if res_conv and res_conv["arquivo"] == nome_arquivo:
            st.success(
                "Excel pronto. "
                + str(res_conv["tabelas"])
                + " tabela(s) identificada(s)."
            )

            st.download_button(
                label="⬇️ Baixar Excel convertido",
                data=res_conv["xlsx"],
                file_name=res_conv["nome_saida"],
                mime=MIME_XLSX,
            )


# ==========================================================
# TELA EXCEL — SEM TÍTULO DE COLUNA
# ==========================================================

elif extensao == ".xlsx":
    st.subheader("Excel → Excel")

    st.info(
        "Não é necessário haver título de coluna. "
        "Escolha a aba, a primeira linha de valores "
        "e uma ou mais colunas."
    )

    st.warning(
        "Células com fórmula (ex.: =E80/3) são mantidas como estão e "
        "NÃO são multiplicadas. Se a fórmula aponta para uma coluna que "
        "você não selecionou, esse preço não será atualizado. "
        "Gráficos e alguns formatos avançados podem se perder ao salvar."
    )

    try:
        workbook_preview = load_workbook(
            io.BytesIO(conteudo_arquivo),
            data_only=False,
            keep_links=True,
            rich_text=True,
        )

        nome_aba = st.selectbox(
            "1. Escolha a aba",
            workbook_preview.sheetnames,
        )

        planilha_preview = workbook_preview[nome_aba]

        linha_inicio = st.number_input(
            "2. Linha onde começam os preços",
            min_value=1,
            max_value=max(1, planilha_preview.max_row),
            value=1,
            step=1,
            help=(
                "Exemplo: se o primeiro preço estiver "
                "em E5, informe 5."
            ),
        )

        letras_colunas = [
            get_column_letter(numero_coluna)
            for numero_coluna in range(
                1,
                planilha_preview.max_column + 1,
            )
        ]

        letras_colunas_selecionadas = st.multiselect(
            "3. Escolha as colunas que deseja alterar",
            letras_colunas,
            default=(["E"] if "E" in letras_colunas else []),
        )

        if not letras_colunas_selecionadas:
            st.warning("Selecione pelo menos uma coluna.")
            st.stop()

        # Converte letras para números
        numeros_colunas = []

        for letra in letras_colunas_selecionadas:
            numero_coluna = 0

            for caractere in letra:
                numero_coluna = (
                    numero_coluna * 26
                    + ord(caractere)
                    - ord("A")
                    + 1
                )

            numeros_colunas.append(numero_coluna)

        st.caption(
            "Colunas selecionadas: "
            + ", ".join(letras_colunas_selecionadas)
            + ". Serão analisadas a partir da linha "
            + str(linha_inicio)
            + "."
        )

        # ==================================================
        # PRÉVIA (mesma regra da geração real)
        # ==================================================

        dados_previa = []

        for letra_coluna, numero_coluna in zip(
            letras_colunas_selecionadas,
            numeros_colunas,
        ):
            for numero_linha in range(
                linha_inicio,
                min(
                    planilha_preview.max_row + 1,
                    linha_inicio + 10,
                ),
            ):
                valor_original = planilha_preview.cell(
                    numero_linha,
                    numero_coluna,
                ).value

                if valor_original is None:
                    continue

                if e_formula(valor_original):
                    novo_valor = manter_formula_original(valor_original)

                elif not parece_preco_excel(valor_original):
                    novo_valor = "Não será alterado"

                else:
                    valor_decimal = converter_valor_excel_para_decimal(
                        valor_original
                    )

                    if valor_decimal is None:
                        novo_valor = "Não será alterado"
                    else:
                        novo_valor = arredondar_sempre_para_cima(
                            valor_decimal * Decimal(str(multiplicador))
                        )

                dados_previa.append(
                    {
                        "Coluna": letra_coluna,
                        "Linha": numero_linha,
                        "Célula": letra_coluna + str(numero_linha),
                        "Valor original": str(valor_original),
                        "Novo valor": str(novo_valor),
                    }
                )

        if dados_previa:
            st.subheader("Prévia")

            st.dataframe(
                pd.DataFrame(dados_previa),
                use_container_width=True,
                hide_index=True,
            )

        else:
            st.warning(
                "Nenhum valor foi encontrado "
                "nas colunas selecionadas."
            )

        # ==================================================
        # GERAR EXCEL
        # ==================================================

        if st.button(
            "Gerar Excel atualizado",
            type="primary",
        ):
            with st.spinner(
                "Atualizando colunas "
                + ", ".join(letras_colunas_selecionadas)
                + "..."
            ):
                resultado, conferencias = atualizar_excel(
                    arquivo_excel=conteudo_arquivo,
                    nome_aba=nome_aba,
                    linha_inicio=linha_inicio,
                    numeros_colunas=numeros_colunas,
                    multiplicador=multiplicador,
                )

            st.session_state["res_xlsx"] = {
                "arquivo": nome_arquivo,
                "xlsx": resultado,
                "conferencias": conferencias,
                "nome_saida": Path(nome_arquivo).stem + "_precos_atualizados.xlsx",
            }

        res_xlsx = st.session_state.get("res_xlsx")

        if res_xlsx and res_xlsx["arquivo"] == nome_arquivo:
            if not res_xlsx["conferencias"]:
                st.warning(
                    "Nenhuma célula foi alterada. "
                    "Verifique as colunas e a linha inicial."
                )
            else:
                st.success(
                    "Excel pronto. "
                    + str(len(res_xlsx["conferencias"]))
                    + " célula(s) processada(s)."
                )

                st.subheader("Conferência")

                st.dataframe(
                    pd.DataFrame(res_xlsx["conferencias"]).astype(str),
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    label="⬇️ Baixar Excel atualizado",
                    data=res_xlsx["xlsx"],
                    file_name=res_xlsx["nome_saida"],
                    mime=MIME_XLSX,
                )

    except Exception as erro:
        st.error(
            "Não foi possível abrir ou salvar este Excel. "
            "Detalhe técnico: " + str(erro)
        )
