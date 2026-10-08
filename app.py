import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Gerador de Overpack - Shipper", layout="centered")

def formatar_numero(valor: float) -> str:
    """Formata um valor flutuante com 2 casas decimais e vírgula como separador."""
    return f"{valor:.2f}".replace('.', ',')

def calcular_texto_overpack(num_boxes: int, weight_per_box: float, overpack_num: int = 1) -> str:
    """
    Realiza o cálculo exato da Shipper Manual:
    Multiplica o número de caixas pelo peso unitário e gera o texto do Overpack.
    """
    total_weight = round(num_boxes * weight_per_box, 2)
    
    str_weight_per_box = formatar_numero(weight_per_box)
    str_total_weight = formatar_numero(total_weight)
    
    linhas = [
        f"{num_boxes} FIBREBOARD BOXES X {str_weight_per_box} Kg G",
        f"OVERPACK USED x {overpack_num}",
        f"#{overpack_num}",
        f"TOTAL QUANTITY PER OVERPACK {str_total_weight} Kg G"
    ]
    
    return "\n\n".join(linhas)

# --- INTERFACE STREAMLIT ---
st.title("📦 Gerador de Declaração de Overpack")
st.write("Ajustado conforme a regra manual do Shipper (Quantidade x Peso Unitário).")

col1, col2, col3 = st.columns(3)

with col1:
    num_boxes = st.number_input("Quantidade de Caixas (Boxes)", min_value=1, value=7, step=1)

with col2:
    weight_per_box = st.number_input("Peso por Caixa (Kg G)", min_value=0.01, value=1.71, step=0.01, format="%.2f")

with col3:
    overpack_num = st.number_input("Número do Overpack", min_value=1, value=1, step=1)

# Cálculo automático
total_calculado = round(num_boxes * weight_per_box, 2)
texto_resultado = calcular_texto_overpack(num_boxes, weight_per_box, overpack_num)

st.subheader("📊 Resumo do Cálculo")
st.info(f"**Cálculo:** {num_boxes} caixas × {formatar_numero(weight_per_box)} kg = **{formatar_numero(total_calculado)} kg**")

st.subheader("📝 Texto Final Gerado")
st.code(texto_
