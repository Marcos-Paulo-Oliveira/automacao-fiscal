import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(page_title="PPC - Consolidado Fiscal", page_icon="📝", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; background-color: #002060; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { width: 100%; background-color: #28a745; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📝 Gerador de Relatório Consolidado")
st.info("Arraste sua **Memória de Cálculo (Base Ouro)** abaixo para gerar o relatório da DCTFWeb.")

# Upload do arquivo gerado anteriormente
arquivo_upload = st.file_uploader("Selecione a Memória de Cálculo (xlsx)", type=["xlsx"])

if arquivo_upload:
    try:
        xl = pd.ExcelFile(arquivo_upload)
        
        # Busca inteligente de dados na primeira aba
        df_busca = xl.parse(xl.sheet_names[0], header=None).astype(str)
        razao_social, cnpj, competencia = "Não Encontrado", "Não Encontrado", "00/0000"

        for i in range(min(len(df_busca), 10)):
            for j in range(len(df_busca.columns)):
                celula = df_busca.iloc[i, j]
                if "RAZÃO SOCIAL:" in celula: razao_social = celula.replace("RAZÃO SOCIAL:", "").strip()
                elif "CNPJ:" in celula: cnpj = celula.replace("CNPJ:", "").strip()
                elif "COMPETÊNCIA" in celula: competencia = celula.split("COMPETÊNCIA")[-1].strip()

        st.success(f"📌 Empresa Detectada: **{razao_social}**")

        # Inputs manuais na lateral ou colunas
        st.subheader("🔢 Informações Manuais (Folha/Trabalho)")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: inss_folha = st.number_input("INSS Folha (eSocial)", min_value=0.0, step=0.01, format="%.2f")
        with col_m2: ir_0588 = st.number_input("IRRF 0588 (Trabalho sem Vínculo)", min_value=0.0, step=0.01, format="%.2f")
        with col_m3: ir_0561 = st.number_input("IRRF 0561 (Trabalho Assalariado)", min_value=0.0, step=0.01, format="%.2f")

        # Função para capturar totais das abas
        def obter_total(nome_aba):
            if nome_aba in xl.sheet_names:
                df_aba = xl.parse(nome_aba)
                if not df_aba.empty:
                    return pd.to_numeric(df_aba.iloc[:, -1], errors='coerce').sum()
            return 0.0

        valores = {
            "1708": obter_total("IRRF 1708"),
            "8045": obter_total("IRRF 8045"),
            "5952": obter_total("CSRF"),
            "1162": obter_total("INSS"),
            "3208": obter_total("IRRF 3208") if "IRRF 3208" in xl.sheet_names else 0.0
        }

        if st.button("🚀 Gerar Relatório Consolidado"):
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "Consolidado Mensal"
            ws.sheet_view.showGridLines = False

            # Estilos
            fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
            font_branca = Font(color='FFFFFF', bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Layout do Relatório
            ws.cell(row=2, column=2, value="DCTFWeb - Relatório Mensal de Impostos Federais Consolidados").font = Font(bold=True, size=14)
            info = [["Razão social", razao_social], ["CNPJ", cnpj], ["Período de apuração", competencia], ["Responsável", "Marcos Paulo Santos de Oliveira"]]
            for i, (label, val) in enumerate(info, 4):
                ws.cell(row=i, column=2, value=label).font = Font(bold=True)
                ws.cell(row=i, column=5, value=val)

            headers = ["Tipo", "Código Retenção RFB", "Valor Retenção", "Descrição", "Observações"]
            for c, h in enumerate(headers, 2):
                cell = ws.cell(row=9, column=c, value=h)
                cell.fill = fill_azul; cell.font = font_branca; cell.border = thin_border; cell.alignment = Alignment(horizontal='center')

            dados_finais = [
                ["INSS", "Folha", inss_folha, "Informação transmitida via eSocial", "Evidência do RH"],
                ["IRRF", "0588", ir_0588, "Rendimento do Trabalho sem Vínculo", ""],
                ["IRRF", "0561", ir_0561, "Rendimento do Trabalho Assalariado", ""],
                ["INSS", "1162", valores["1162"], "Retenção na fonte NFSe (EFD REINF)", "Memória de Cálculo"],
                ["IRRF", "1708", valores["1708"], "Serviços Prestados PJ", ""],
                ["IRRF", "8045", valores["8045"], "Outros Rendimentos", ""],
                ["IRRF", "3208", valores["3208"], "Aluguéis PF", ""],
                ["CSRF", "5952", valores["5952"], "Retenção CSRF (PCC)", ""]
            ]

            for r_idx, linha in enumerate(dados_finais, 10):
                for c_idx, valor in enumerate(linha, 2):
                    cell = ws.cell(row=r_idx, column=c_idx, value=valor)
                    cell.border = thin_border
                    if c_idx == 4: cell.number_format = 'R$ #,##0.00'

            for i, largura in enumerate([12, 18, 20, 45, 30], 2):
                ws.column_dimensions[get_column_letter(i)].width = largura

            wb.save(output)
            st.download_button(
                label="📥 Baixar Relatório Consolidado",
                data=output.getvalue(),
                file_name=f"Consolidado - {razao_social} - {competencia}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
