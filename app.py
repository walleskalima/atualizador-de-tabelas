import io
import re
from decimal import Decimal, InvalidOperation, ROUND_CEILING
from pathlib import Path

import pandas as pd
import pymupdf
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================

st.set_page_config(
    page_title="Atualizador de Preços",
    page_icon="💰",
    layout="wide",
)

st.title("💰 Atualizador de Preços")
st.caption(
    "PDF entra como PDF e sai como PDF. "
    "Excel entra como Excel e sai como Excel."
)


# ==========================================================
# PREÇOS, MULTIPLICAÇÃO E ARREDONDAMENTO
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
    Faz:
    novo valor = teto(valor original x multiplicador)
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
    """
    Converte cor em inteiro RGB do PyMuPDF para tupla RGB de 0 a 1.
    """
    if not isinstance(cor, int):
        return (0, 0, 0)

    vermelho = ((cor >> 16) & 255) / 255
    verde = ((cor >> 8) & 255) / 255
    azul = (cor & 255) / 255

    return (vermelho, verde, azul)


def escolher_fonte_pdf(nome_fonte_original):
    """
    Escolhe uma fonte-base segura para reescrever o preço.
    """
    nome = str(nome_fonte_original).lower()

    if "bold" in nome:
        return "hebo"

    if "times" in nome:
        return "tiro"

    if "courier" in nome:
        return "cour"

    return "helv"


def medir_largura_texto(texto, fonte, tamanho):
    """
    Mede largura aproximada do texto novo para caber no espaço do preço.
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
    Encontra spans de texto que possuem preços como:
    R$ 926,82
    R$ 1.073,88
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

                correspondencias = list(
                    PADRAO_PRECO.finditer(texto_span)
                )

                if not correspondencias:
                    continue

                x0, y0, x1, y1 = span["bbox"]
                largura_span = x1 - x0
                tamanho_texto = max(len(texto_span), 1)

                for correspondencia in correspondencias:
                    inicio = correspondencia.start()
                    fim = correspondencia.end()

                    # Quando o span contém apenas o preço,
                    # usa a área exata do próprio span.
                    if inicio == 0 and fim == len(texto_span):
                        caixa_preco = pymupdf.Rect(
                            x0,
                            y0,
                            x1,
                            y1,
                        )

                    # Se houver outro texto junto do preço,
                    # calcula uma faixa proporcional dentro do span.
                    else:
                        x_inicio = x0 + (
                            largura_span * inicio / tamanho_texto
                        )

                        x_fim = x0 + (
                            largura_span * fim / tamanho_texto
                        )

                        caixa_preco = pymupdf.Rect(
                            x_inicio,
                            y0,
                            x_fim,
                            y1,
                        )

                    encontrados.append(
                        {
                            "texto_original": correspondencia.group(0),
                            "rect": caixa_preco,
                            "fonte_original": span.get("font", "helv"),
                            "tamanho_fonte": span.get("size", 10),
                            "cor": rgb_inteiro_para_tupla(
                                span.get("color", 0)
                            ),
                        }
                    )

    return encontrados


def colocar_fundo_branco(pagina):
    """
    Coloca um retângulo branco como plano de fundo da página.

    overlay=False:
    - coloca atrás dos textos, fotos e imagens existentes;
    - mantém o restante do catálogo visível;
    - funciona para páginas cujo fundo não seja uma imagem/gráfico
      por cima de tudo.

    Não use overlay=True aqui: ele cobriria fotos, logos e textos.
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
    Recebe bytes de PDF e devolve bytes de outro PDF.

    Mantém o catálogo em PDF e altera somente os preços encontrados.
    """
    documento = pymupdf.open(
        stream=arquivo_pdf,
        filetype="pdf",
    )

    conferencias = []

    # Primeiro: cria fundo branco atrás do conteúdo da página.
    # Este loop precisa ficar FORA do loop que troca preços.
    if deixar_fundo_branco:
        for pagina in documento:
            colocar_fundo_branco(pagina)

    # Depois: localiza e altera preços página por página.
    for numero_pagina, pagina in enumerate(documento, start=1):
        precos_pagina = localizar_precos_na_pagina(pagina)
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

            # Pequena margem para remover completamente
            # os caracteres do preço antigo.
            caixa_redacao = pymupdf.Rect(
                caixa_original.x0 - 1.5,
                caixa_original.y0 - 1.0,
                caixa_original.x1 + 1.5,
                caixa_original.y1 + 1.0,
            )

            # Branco somente atrás do preço que será alterado.
            pagina.add_redact_annot(
                caixa_redacao,
                fill=(1, 1, 1),
            )

            substituicoes.append(
                {
                    "texto_original": texto_original,
                    "texto_novo": texto_novo,
                    "caixa_original": caixa_original,
                    "fonte": escolher_fonte_pdf(
                        preco["fonte_original"]
                    ),
                    "tamanho": preco["tamanho_fonte"],
                    "cor": preco["cor"],
                }
            )

        # Remove os preços antigos antes de escrever os novos.
        if substituicoes:
            pagina.apply_redactions(
                images=pymupdf.PDF_REDACT_IMAGE_NONE,
                graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                text=pymupdf.PDF_REDACT_TEXT_REMOVE,
            )

        # Escreve os preços novos nas posições originais.
        for item in substituicoes:
            caixa = item["caixa_original"]
            texto_novo = item["texto_novo"]
            fonte = item["fonte"]
            cor = item["cor"]

            tamanho_original = item["tamanho"]
            tamanho_ajustado = tamanho_original

            largura_original = caixa.width
            largura_nova = medir_largura_texto(
                texto_novo,
                fonte,
                tamanho_ajustado,
            )

            # Reduz a fonte apenas se o novo preço não couber.
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

            # Caixa onde o novo texto será inserido.
            caixa_texto = pymupdf.Rect(
                caixa.x0,
                caixa.y0 - 0.5,
                max(caixa.x1, caixa.x0 + largura_nova + 3),
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

            # Se não couber pelo textbox, escreve diretamente.
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
                    "Página": numero_pagina,
                    "Preço original": item["texto_original"],
                    "Preço atualizado": item["texto_novo"],
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

    return arquivo_final.getvalue(), conferencias


# ==========================================================
# EXCEL -> EXCEL
# ==========================================================

def converter_valor_excel_para_decimal(valor):
    """
    Converte número ou texto monetário para Decimal.
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
    Aceita valores numéricos ou texto no padrão monetário.
    Evita mudar códigos, medidas e descrições.
    """
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


def obter_cabecalhos(planilha, linha_cabecalho):
    """
    Lê os títulos das colunas da linha escolhida pelo usuário.
    """
    cabecalhos = {}

    for numero_coluna in range(1, planilha.max_column + 1):
        valor = planilha.cell(
            linha_cabecalho,
            numero_coluna,
        ).value

        if valor is None or not str(valor).strip():
            continue

        nome = str(valor).strip()

        if nome in cabecalhos:
            nome = f"{nome} ({numero_coluna})"

        cabecalhos[nome] = numero_coluna

    return cabecalhos


def atualizar_excel(
    arquivo_excel,
    nome_aba,
    linha_cabecalho,
    numero_coluna,
    multiplicador,
):
    """
    Altera somente as células da coluna selecionada.
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

        # Não altera fórmulas.
        if (
            isinstance(valor_original, str)
            and valor_original.startswith("=")
        ):
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

        # Mantém formato textual R$ se já era texto com R$.
        if (
            isinstance(valor_original, str)
            and "R$" in valor_original.upper()
        ):
            celula.value = formatar_preco_brasileiro(
                novo_valor
            )

        # Mantém como número editável no Excel.
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
            }
        )

    arquivo_final = io.BytesIO()
    workbook.save(arquivo_final)
    arquivo_final.seek(0)

    return arquivo_final.getvalue(), conferencias


# ==========================================================
# INTERFACE DO APLICATIVO
# ==========================================================

arquivo = st.file_uploader(
    "Envie um arquivo PDF ou Excel",
    type=["pdf", "xlsx"],
)

multiplicador = st.number_input(
    "Valor para multiplicar os preços",
    min_value=0.000001,
    value=2.383949988,
    step=0.000001,
    format="%.9f",
)

if arquivo is None:
    st.info("Envie um PDF ou Excel para começar.")
    st.stop()

nome_arquivo = arquivo.name
extensao = Path(nome_arquivo).suffix.lower()
conteudo_arquivo = arquivo.getvalue()


# ==========================================================
# TELA PARA PDF
# ==========================================================

if extensao == ".pdf":
    st.subheader("PDF → PDF")

    deixar_fundo_branco = st.checkbox(
        "Adicionar fundo branco atrás do conteúdo da página",
        value=True,
        help=(
            "Mantém textos e fotos existentes, pois a camada branca é "
            "colocada atrás deles. Não remove imagens coloridas que já "
            "estejam desenhadas por cima do fundo."
        ),
    )

    st.info(
        "O PDF final mantém as páginas e a estrutura original. "
        "Somente os preços no padrão R$ 0,00 serão atualizados."
    )

    if st.button(
        "Gerar PDF com preços atualizados",
        type="primary",
    ):
        try:
            with st.spinner(
                "Localizando preços e criando o novo PDF..."
            ):
                resultado, conferencias = atualizar_pdf_mantendo_layout(
                    arquivo_pdf=conteudo_arquivo,
                    multiplicador=multiplicador,
                    deixar_fundo_branco=deixar_fundo_branco,
                )

            if not conferencias:
                st.warning(
                    "Nenhum preço foi encontrado. "
                    "O PDF pode ser escaneado como imagem ou usar outro formato."
                )
                st.stop()

            st.success(
                f"PDF pronto. {len(conferencias)} preço(s) foram atualizados."
            )

            st.subheader("Conferência dos preços alterados")
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
                "Não foi possível processar este PDF. "
                f"Detalhe técnico: {erro}"
            )


# ==========================================================
# TELA PARA EXCEL
# ==========================================================

elif extensao == ".xlsx":
    st.subheader("Excel → Excel")

    st.info(
        "Selecione a aba e a coluna de preços. "
        "O sistema altera somente os valores da coluna escolhida."
    )

    try:
        workbook_preview = load_workbook(
            io.BytesIO(conteudo_arquivo),
            data_only=False,
            keep_links=True,
            rich_text=True,
        )

        nome_aba = st.selectbox(
            "1. Escolha a aba da tabela",
            workbook_preview.sheetnames,
        )

        planilha_preview = workbook_preview[nome_aba]

        linha_cabecalho = st.number_input(
            "2. Em qual linha estão os títulos das colunas?",
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
                "Tente outro número de linha."
            )
            st.stop()

        nome_coluna = st.selectbox(
            "3. Escolha a coluna de preço/valor",
            list(cabecalhos.keys()),
        )

        numero_coluna = cabecalhos[nome_coluna]

        st.caption(
            f"Coluna selecionada: {nome_coluna} "
            f"({get_column_letter(numero_coluna)})."
        )

        dados_previa = []

        for numero_linha in range(
            linha_cabecalho + 1,
            min(
                planilha_preview.max_row + 1,
                linha_cabecalho + 11,
            ),
        ):
            valor = planilha_preview.cell(
                numero_linha,
                numero_coluna,
            ).value

            if valor is None:
                continue

            valor_decimal = converter_valor_excel_para_decimal(
                valor
            )

            if valor_decimal is None:
                novo_valor = "Não será alterado"
            else:
                novo_valor = arredondar_sempre_para_cima(
                    valor_decimal * Decimal(str(multiplicador))
                )

            dados_previa.append(
                {
                    "Linha": numero_linha,
                    "Valor original": valor,
                    "Novo valor": novo_valor,
                }
            )

        if dados_previa:
            st.subheader("Prévia")
            st.dataframe(
                pd.DataFrame(dados_previa),
                use_container_width=True,
                hide_index=True,
            )

        if st.button(
            "Gerar Excel atualizado",
            type="primary",
        ):
            with st.spinner(
                "Atualizando a coluna de preços..."
            ):
                resultado, conferencias = atualizar_excel(
                    arquivo_excel=conteudo_arquivo,
                    nome_aba=nome_aba,
                    linha_cabecalho=linha_cabecalho,
                    numero_coluna=numero_coluna,
                    multiplicador=multiplicador,
                )

            if not conferencias:
                st.warning(
                    "Nenhum valor foi alterado. Verifique se a coluna escolhida "
                    "contém números ou preços no formato R$ 0,00."
                )
                st.stop()

            st.success(
                f"Excel pronto. {len(conferencias)} valor(es) foram atualizados."
            )

            st.subheader("Conferência dos valores alterados")
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
            "Não foi possível abrir ou salvar este Excel. "
            f"Detalhe técnico: {erro}"
        )
