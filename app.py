import io
import math
import os
from datetime import datetime
from decimal import Decimal
import pandas as pd
import pytz
import openpyxl
import streamlit as st

# ReportLab para geração de PDF e tratamento de imagens
from reportlab.lib.pagesizes import A4, portrait
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Configuração Inicial da Página
st.set_page_config(
    page_title="Sistema Unificado de Embarque - New Post",
    page_icon="📦",
    layout="wide"
)

# Constantes e Mapeamentos
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE",
    "CGB": "CUIABA",
    "CWB": "CURITIBA",
    "FLN": "FLORIANOPOLIS",
    "GYN": "GOIANIA",
    "MAO": "MANAUS",
    "POA": "PORTO ALEGRE",
    "PVH": "PORTO VELHO",
    "POA PRIME": "PRIME-RS PORTO ALEGRE",
    "FLN PRIME": "PRIME-SC FLORIANÓPOLIS",
}

TODAS_SIGLAS = [
    "CGR", "CGB", "CWB", "FLN", "GYN", "MAO", "POA", "PVH", "POA PRIME", "FLN PRIME"
]

# ---------------------------------------------------------
# FUNÇÕES DE APOIO - CÁLCULOS E EXTRATOR
# ---------------------------------------------------------
def extrair_dados_coleta(df_raw, termo_busca):
    linha_cabecalho = None
    idx_destino, idx_qntde, idx_peso = None, None, None

    for idx, row in df_raw.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notnull(v)]
        if "DESTINO" in valores and ("QNTDE" in valores or "QNTD" in valores) and "PESO" in valores:
            linha_cabecalho = idx
            valores_linha_lista = [str(v).strip().upper() for v in row.values]
            idx_destino = valores_linha_lista.index("DESTINO")
            idx_qntde = (
                valores_linha_lista.index("QNTDE")
                if "QNTDE" in valores_linha_lista
                else valores_linha_lista.index("QNTD")
            )
            idx_peso = valores_linha_lista.index("PESO")
            break

    if linha_cabecalho is None:
        return None, None

    for idx in range(linha_cabecalho + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        val_destino = (
            str(row.iloc[idx_destino]).strip().upper()
            if pd.notnull(row.iloc[idx_destino])
            else ""
        )

        if "PRIME" in termo_busca:
            if termo_busca not in val_destino:
                continue
        else:
            if "TOTAL" in val_destino or val_destino == "" or val_destino.isdigit():
                continue
            destino_limpo = (
                val_destino.replace("AGF", "")
                .replace(" MT", "")
                .replace(" MS", "")
                .replace(" PR", "")
                .replace(" SC", "")
                .replace(" GO", "")
                .replace(" AM", "")
                .replace(" RS", "")
                .replace(" RO", "")
                .strip()
            )
            if termo_busca not in destino_limpo and destino_limpo not in termo_busca:
                continue

        try:
            val_q = row.iloc[idx_qntde]
            qtd_volumes = int(float(str(val_q).replace(",", ".").strip()))
            val_p = row.iloc[idx_peso]
            peso_original = float(str(val_p).replace(",", ".").strip())
            return qtd_volumes, peso_original
        except Exception:
            continue
    return None, None


def calcular_peso_total(f_sacas, q_volumes, p_original):
    f_sacas_dec = Decimal(str(f_sacas))
    d_peso_original = Decimal(str(p_original))

    g_peso_corrigido = (f_sacas_dec * Decimal("3")) + d_peso_original

    fracao_fib = float(q_volumes) / float(f_sacas)
    i_fib = Decimal(
        str(
            max(
                1,
                math.floor(fracao_fib)
                + (1 if (fracao_fib - math.floor(fracao_fib)) >= 0.5 else 0),
            )
        )
    )

    base_j = float(g_peso_corrigido / f_sacas_dec / i_fib)
    j_inicio = Decimal(f"{max(0.01, math.floor(base_j * 100) / 100 - 0.50):.2f}")
    perfeito_j = j_inicio

    for a in range(1500):
        j_teste = j_inicio + (Decimal(str(a)) * Decimal("0.01"))
        if (j_teste * i_fib * f_sacas_dec) - g_peso_corrigido >= 0:
            perfeito_j = j_teste
            break

    total_overpack = perfeito_j * i_fib
    peso_total_destino = total_overpack * f_sacas_dec
    return float(peso_total_destino)


def gerar_excel_controle_embarque(cia, data_str, dados_linhas, caminhao_str, condutor_str, template_path="Controle Embarque-t.xlsx"):
    if not os.path.exists(template_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Controle Embarque Aeroporto"
    else:
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

    ws.cell(row=2, column=9, value=data_str)
    ws.cell(row=5, column=5, value=cia.upper())

    linhas_processadas = set()
    for row in range(7, 20):
        val_sigla = str(ws.cell(row=row, column=1).value or "").strip().upper()
        if val_sigla in dados_linhas and val_sigla not in linhas_processadas:
            linhas_processadas.add(val_sigla)
            info = dados_linhas[val_sigla]
            
            if info["sacas"] != "":
                ws.cell(row=row, column=2, value=int(info["sacas"]))
            else:
                ws.cell(row=row, column=2, value="")
                
            if isinstance(info["peso"], (int, float)):
                ws.cell(row=row, column=8, value=round(info["peso"], 2))
            else:
                ws.cell(row=row, column=8, value="")

    ws.cell(row=19, column=3, value=caminhao_str.upper())
    ws.cell(row=20, column=3, value=condutor_str.upper())

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_pdf_controle_embarque(cia, data_str, dados_linhas, caminhao_str, condutor_str):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=portrait(A4))
    width, height = A4

    possiveis_nomes = ["logo.png.JPG", "logo.png.jpg", "logo.png", "logo.jpg"]
    path_logo = None
    for nome in possiveis_nomes:
        if os.path.exists(nome):
            path_logo = nome
            break

    if path_logo:
        try:
            img = ImageReader(path_logo)
            c.drawImage(img, 40, height - 75, width=120, height=50, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2.0, height - 50, "CONTROLE EMBARQUE")

    box_date_w, box_date_h = 140, 24
    box_date_x = width - 40 - box_date_w
    box_date_y = height - 58
    c.setLineWidth(1.5)
    c.rect(box_date_x, box_date_y, box_date_w, box_date_h)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(box_date_x + 8, box_date_y + 7, f"DATA:  {data_str}")

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2.0, height - 105, cia.upper())
    c.setLineWidth(1)
    cia_text_w = c.stringWidth(cia.upper(), "Helvetica-Bold", 18)
    c.line((width - cia_text_w) / 2.0, height - 108, (width + cia_text_w) / 2.0, height - 108)

    y_start = height - 150
    row_height = 28
    box_width = 50

    x_sigla = 70
    x_box1 = 140
    x_lbl1 = x_box1 + box_width + 10
    x_times1 = x_lbl1 + 55

    x_box2 = x_times1 + 25
    x_lbl2 = x_box2 + box_width + 10
    x_times2 = x_lbl2 + 65

    x_box3 = x_times2 + 25
    x_lbl3 = x_box3 + box_width + 10

    y_curr = y_start

    for sigla in TODAS_SIGLAS:
        info = dados_linhas.get(sigla, {"sacas": "", "peso": ""})
        qnt_sacas = str(info["sacas"]) if info["sacas"] != "" else ""
        peso_total = f"{info['peso']:.2f}".replace('.', ',') if isinstance(info["peso"], (int, float)) else ""

        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_sigla, y_curr + 6, sigla)

        c.setLineWidth(1.2)
        c.rect(x_box1, y_curr, box_width, row_height - 6)
        if qnt_sacas:
            c.setFont("Helvetica", 10)
            c.drawCentredString(x_box1 + (box_width / 2), y_curr + 6, qnt_sacas)
        c.setFont("Helvetica", 10)
        c.drawString(x_lbl1, y_curr + 6, "SACAS")
        c.drawString(x_times1, y_curr + 6, "X")

        c.rect(x_box2, y_curr, box_width, row_height - 6)
        c.drawString(x_lbl2, y_curr + 6, "PALETES")
        c.drawString(x_times2, y_curr + 6, "X")

        c.rect(x_box3, y_curr, box_width, row_height - 6)
        if peso_total:
            c.setFont("Helvetica", 9)
            c.drawCentredString(x_box3 + (box_width / 2), y_curr + 6, peso_total)
        c.setFont("Helvetica", 10)
        c.drawString(x_lbl3, y_curr + 6, "PESO")

        y_curr -= row_height

    y_footer = y_curr - 30
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_sigla + 30, y_footer, f"CAMINHÃO {caminhao_str.upper()}")
    c.drawString(x_sigla + 30, y_footer - 20, f"CONDUTOR: {condutor_str.upper()}")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def gerar_shippers_gerais(df_raw, dados_embarque, cia, data_str):
    """
    Função para processar e gerar os Shippers com base nos dados únicos informados.
    Pode retornar um PDF unificado, ZIP com documentos, etc.
    """
    # Exemplo simples de buffer para integrar a lógica do seu código de Shippers
    buffer_shippers = io.BytesIO()
    buffer_shippers.write(b"SHIPPERS GERADOS COM SUCESSO")
    buffer_shippers.seek(0)
    return buffer_shippers.getvalue()


# =========================================================
# INTERFACE PRINCIPAL - ENTRADA ÚNICA DE DADOS
# =========================================================
st.title("🚚 Gerador Unificado: Shippers & Controle de Embarque")
st.markdown("Preencha os dados abaixo apenas **uma vez** para gerar **ambos os documentos simultaneamente**.")

# 1. Informações Gerais do Embarque
st.subheader("1. Informações do Embarque")
col1, col2, col3 = st.columns(3)
with col1:
    cia_input = st.text_input("Companhia / Título Central:", value="LATAM")
with col2:
    caminhao_input = st.text_input("Identificação do Caminhão:", value="1º")
with col3:
    condutor_input = st.text_input("Nome do Condutor:", value="ANTONIO")

# 2. Siglas e Sacas
st.markdown("---")
st.subheader("2. Destinos e Quantidade de Sacas")

siglas_input = st.text_input(
    "Siglas do embarque (separadas por vírgula):",
    value="CGR, CGB, CWB, FLN, GYN, MAO, POA, PVH, POA PRIME, FLN PRIME",
).upper().strip()

lista_siglas_usuario = [s.strip() for s in siglas_input.split(",") if s.strip()]

sacas_manuais = {}
if lista_siglas_usuario:
    cols = st.columns(min(len(lista_siglas_usuario), 5))
    for idx, sigla in enumerate(lista_siglas_usuario):
        with cols[idx % 5]:
            sacas_manuais[sigla] = st.number_input(
                f"Sacas {sigla}:",
                min_value=1,
                value=None,
                step=1,
                key=f"sacas_{sigla}",
            )

# 3. Planilha de Coleta
st.markdown("---")
st.subheader("3. Carregue a Planilha de Coleta")
file_excel = st.file_uploader("Envie a planilha de coleta (.xlsx / .xlsm):", type=["xlsx", "xlsm"])

todas_preenchidas = len(sacas_manuais) > 0 and all(v is not None for v in sacas_manuais.values())

# 4. Processamento Simultâneo
st.markdown("---")

if file_excel and todas_preenchidas:
    if st.button("🚀 GERAR TUDO SIMULTANEAMENTE", use_container_width=True):
        try:
            df_raw = pd.read_excel(file_excel, header=None, engine="openpyxl")
            dados_embarque = {}

            # Processa o cálculo para cada sigla
            for sigla in TODAS_SIGLAS:
                if sigla in sacas_manuais and sacas_manuais[sigla] is not None:
                    qnt_sacas = sacas_manuais[sigla]
                    cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                    q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                    if p_original and q_volumes:
                        peso_calc = calcular_peso_total(qnt_sacas, q_volumes, p_original)
                        dados_embarque[sigla] = {"sacas": qnt_sacas, "peso": peso_calc}
                    else:
                        dados_embarque[sigla] = {"sacas": qnt_sacas, "peso": ""}
                else:
                    dados_embarque[sigla] = {"sacas": "", "peso": ""}

            fuso_sp = pytz.timezone("America/Sao_Paulo")
            data_hoje = datetime.now(fuso_sp).strftime("%d/%m/%Y")
            data_file = datetime.now(fuso_sp).strftime("%Y%m%d")

            # Gerar PDF e Excel do Controle de Embarque
            pdf_emb_bytes = gerar_pdf_controle_embarque(
                cia=cia_input,
                data_str=data_hoje,
                dados_linhas=dados_embarque,
                caminhao_str=caminhao_input,
                condutor_str=condutor_input
            )

            excel_emb_bytes = gerar_excel_controle_embarque(
                cia=cia_input,
                data_str=data_hoje,
                dados_linhas=dados_embarque,
                caminhao_str=caminhao_input,
                condutor_str=condutor_input
            )

            # Gerar os Shippers
            shippers_bytes = gerar_shippers_gerais(
                df_raw=df_raw,
                dados_embarque=dados_embarque,
                cia=cia_input,
                data_str=data_hoje
            )

            st.success("✅ **Todos os documentos foram gerados com sucesso!** Escolha abaixo os arquivos para baixar:")

            # Área de Downloads Dividida em Blocos
            st.markdown("### 📥 Downloads Disponíveis")
            col_down1, col_down2 = st.columns(2)

            with col_down1:
                st.markdown("#### 📄 Documentos de Controle de Embarque")
                st.download_button(
                    label="BAIXAR CONTROLE DE EMBARQUE (PDF)",
                    data=pdf_emb_bytes,
                    file_name=f"Controle_Embarque_{cia_input}_{data_file}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                st.download_button(
                    label="BAIXAR CONTROLE DE EMBARQUE (EXCEL)",
                    data=excel_emb_bytes,
                    file_name=f"Controle_Embarque_{cia_input}_{data_file}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_down2:
                st.markdown("#### 📦 Documentos de Shippers")
                st.download_button(
                    label="BAIXAR SHIPPERS GERADOS",
                    data=shippers_bytes,
                    file_name=f"Shippers_{cia_input}_{data_file}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
