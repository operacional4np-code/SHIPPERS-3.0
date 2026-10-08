import io
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

# Configuração da página Streamlit
st.set_page_config(
    page_title="Gerador de Declaração de Overpack - Shipper",
    page_icon="📦",
    layout="wide"
)

# ==============================================================================
# 1. FUNÇÕES DE CÁLCULO E FORMATAÇÃO (AJUSTADO PARA O SHIPPER MANUAL)
# ==============================================================================

def formatar_peso(valor: float) -> str:
    """Formata valor numérico para o padrão com vírgula (ex: 1,71)."""
    return f"{valor:.2f}".replace('.', ',')

def calcular_overpack(num_boxes: int, weight_per_box: float, overpack_num: int = 1) -> dict:
    """
    Realiza a multiplicação direta (Caixas x Peso Unitário) conforme a regra manual.
    Exemplo: 7 caixas x 1,71 Kg = 11,97 Kg G
    """
    # Cálculo exato ajustado
    total_weight = round(num_boxes * weight_per_box, 2)
    
    str_weight_per_box = formatar_peso(weight_per_box)
    str_total_weight = formatar_peso(total_weight)
    
    # Estrutura do texto idêntica ao espelho da Shipper
    texto_declaracao = (
        f"{num_boxes} FIBREBOARD BOXES X {str_weight_per_box} Kg G\n\n"
        f"OVERPACK USED x {overpack_num}\n\n"
        f"#{overpack_num}\n\n"
        f"TOTAL QUANTITY PER OVERPACK {str_total_weight} Kg G"
    )
    
    return {
        "num_boxes": num_boxes,
        "weight_per_box": weight_per_box,
        "total_weight": total_weight,
        "str_weight_per_box": str_weight_per_box,
        "str_total_weight": str_total_weight,
        "texto_declaracao": texto_declaracao
    }

# ==============================================================================
# 2. GERADOR DE PDF COMPLETO (REPORTLAB)
# ==============================================================================

def gerar_pdf_overpack(dados_overpack: list) -> bytes:
    """Gera um PDF formatado contendo as caixas/etiquetas de Overpack."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    elements = []
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        alignment=1, # Centralizado
        spaceAfter=20
    )
    style_body = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontSize=11,
        leading=16,
        alignment=0 # Esquerda
    )
    
    elements.append(Paragraph("<b>DECLARAÇÃO DE OVERPACK / SHIPPER</b>", style_title))
    elements.append(Spacer(1, 10))
    
    for item in dados_overpack:
        texto_formatado = item['texto_declaracao'].replace('\n', '<br/>')
        p = Paragraph(texto_formatado, style_body)
        
        # Moldura retangular estilo etiqueta
        t = Table([[p]], colWidths=[500])
        t.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('PADDING', (0, 0), (-1, -1), 12),
            ('BACKGROUND', (0, 0), (-1, -1), colors.whitesmoke)
        ]))
        elements.append(t)
        elements.append(Spacer(1, 15))
        
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# 3. INTERFACE STREAMLIT
# ==============================================================================

st.title("📦 Sistema de Geração de Overpack - Ajustado")
st.markdown("Cálculo parametrizado por multiplicação direta: **Total = Caixas × Peso Unitário**.")

aba1, aba2 = st.tabs(["📝 Entrada Manual", "📁 Processar Planilha (Excel/CSV)"])

# ------------------------------------------------------------------------------
# ABA 1: GERAR INDIVIDUALMENTE
# ------------------------------------------------------------------------------
with aba1:
    st.subheader("Parâmetros do Overpack")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        num_boxes = st.number_input("Quantidade de Caixas (Boxes)", min_value=1, value=7, step=1, key="m_boxes")
    with col2:
        weight_per_box = st.number_input("Peso por Caixa (Kg G)", min_value=0.01, value=1.71, step=0.01, format="%.2f", key="m_weight")
    with col3:
        overpack_num = st.number_input("Número do Overpack", min_value=1, value=1, step=1, key="m_num")
        
    res = calcular_overpack(num_boxes, weight_per_box, overpack_num)
    
    st.markdown("---")
    st.subheader("Resumo do Cálculo")
    st.success(f"**Fórmula:** {num_boxes} caixas × {res['str_weight_per_box']} kg = **{res['str_total_weight']} kg G**")
    
    st.subheader("Texto Final Gerado")
    st.code(res["texto_declaracao"], language="text")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📄 Baixar Texto (.txt)",
            data=res["texto_declaracao"],
            file_name=f"overpack_{overpack_num}.txt",
            mime="text/plain",
            use_container_width=True
        )
    with col_d2:
        pdf_bytes = gerar_pdf_overpack([res])
        st.download_button()
        label="🔴 Baixar Documento PDF",
        data=
