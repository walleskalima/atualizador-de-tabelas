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
