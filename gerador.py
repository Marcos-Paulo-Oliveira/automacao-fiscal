import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Configuração da página (Ícone de gráfico e layout mais fluido)
st.set_page_config(page_title="PPC - Automação Fiscal", page_icon="📊", layout="wide")

# Estilização Personalizada (Cores da PPC e ajustes de botões)
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; background-color: #002060; color: white; border-radius: 8px; height: 3em; font-weight: bold; }
    .stDownloadButton>button { width: 100%; background-color: #28a745; color: white; border-radius: 8px; height: 3em; font-weight: bold; }
    .stDownloadButton>button:hover { background-color: #218838; color: white; }
    </style>
    """, unsafe_allow_html=True)

def aplicar_estilo_ppc(writer, df_filtrado, colunas_mapeadas, nome_aba, titulo_imposto, razao, cnpj, comp):
    ws = writer.book.create_sheet(nome_aba)
    writer.sheets[nome_aba] = ws
    ws.sheet_view.showGridLines = False 
    ws.column_dimensions['A'].width = 3

    fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    font_branca = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    # Cabeçalho (Coluna B até J)
    for row_num, texto in enumerate([f'RAZÃO SOCIAL: {razao}', f'CNPJ: {cnpj}', f'{titulo_imposto} - COMPETÊNCIA {comp}'], 2):
        ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=10)
        cell = ws.cell(row=row_num, column=2)
        cell.value = texto
        cell.alignment = align_center
        cell.font = Font(bold=True, size=12)

    for col_num, header in enumerate(colunas_mapeadas.values(), 2):
        cell = ws.cell(row=6, column=col_num)
        cell.value = header
        cell.fill = fill_azul
        cell.font = font_branca
        cell.alignment = align_center
        cell.border = thin_border

    moeda_cols = ['Vlr Contábil', 'Base IRRF', 'Valor IRRF', 'Base CSR', 'Total PCC', 'ISS', 'Valor INSS', 'Base INSS', 'Base ISS']

    if df_filtrado.empty:
        row_msg = 7
        ws.merge_cells(start_row=row_msg, start_column=2, end_row=row_msg, end_column=10)
        cell_msg = ws.cell(row=row_msg, column=2)
        cell_msg.value = "SEM MOVIMENTO"
        cell_msg.font = Font(bold=True)
        cell_msg.alignment = align_center
        for col_idx in range(2, 11):
            ws.cell(row=row_msg, column=col_idx).border = thin_border
    else:
        dados_finais = df_filtrado[list(colunas_mapeadas.keys())].rename(columns=colunas_mapeadas)
        for r_idx, row in enumerate(dados_finais.values, 7):
            for c_idx, value in enumerate(row, 2):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value
                cell.border = thin_border
                header_text = list(colunas_mapeadas.values())[c_idx-2]
                
                if header_text in moeda_cols:
                    cell.number_format = 'R$ #,##0.00'
                elif 'Data' in header_text:
                    cell.number_format = 'dd/mm/yyyy'
                elif 'Aliq' in header_text or '%' in header_text:
                    cell.number_format = '0.00%'

        last_row = 6 + len(dados_finais)
        row_total = last_row + 1
        ws.merge_cells(start_row=row_total, start_column=2, end_row=row_total, end_column=6)
        cell_total_label = ws.cell(row=row_total, column=2)
        cell_total_label.value = "TOTAL"
        cell_total_label.font = Font(bold=True)
        cell_total_label.alignment = Alignment(horizontal='right')
        
        for col_idx in range(2, 7):
            ws.cell(row=row_total, column=col_idx).border = thin_border
        
        for col_idx in range(7, len(colunas_mapeadas) + 2):
            ws.cell(row=row_total, column=col_idx).border = thin_border
            header_text = list(colunas_mapeadas.values())[col_idx-2]
            if header_text in moeda_cols:
                col_letter = get_column_letter(col_idx)
                cell_sum = ws.cell(row=row_total, column=col_idx)
                cell_sum.value = f"=SUM({col_letter}7:{col_letter}{last_row})"
                cell_sum.font = Font(bold=True)
                cell_sum.number_format = 'R$ #,##0.00'

    for col in ws.columns:
        column_letter = col[0].column_letter
        if column_letter == 'A': continue
        max_length = 0
        for cell in col:
            if cell.value:
                length = len(str(cell.value))
                if length > max_length: max_length = length
        ws.column_dimensions[column_letter].width = max_length + 4

# --- INTERFACE STREAMLIT ---
with st.sidebar:
    st.image("https://www.ppcaudit.com.br/wp-content/uploads/2017/07/logo-ppc.png", width=150)
    st.title("⚙️ Painel Fiscal")
    st.markdown("---")
    st.info("Este app transforma o relatório bruto do UneCont em uma Memória de Cálculo formatada.")
    st.caption("Desenvolvido para Marcos Paulo | Versão 2.1")

st.title("📊 Gerador de Memória de Cálculo")
st.markdown("Arraste o arquivo baixado do **UneCont** abaixo para iniciar o processamento.")

arquivo_upload = st.file_uploader("Selecione o arquivo UneCont (xlsx)", type=["xlsx"])

if arquivo_upload:
    try:
        df = pd.read_excel(arquivo_upload)
        
        # --- CORREÇÃO DO CÓDIGO DE SERVIÇO (Troca , por .) ---
        if 'Serviço Federal' in df.columns:
            df['Serviço Federal'] = df['Serviço Federal'].astype(str).str.replace(',', '.', regex=False)

        df['ISS_TOTAL'] = df['ISS Dentro do Município'].fillna(0) + df['ISS Fora do Município'].fillna(0)
        df['BASE_ISS_TOTAL'] = df['Base de Cálculo ISS'].fillna(0)
        df['ALIQ_ISS_TOTAL'] = df['% ISS Dentro do Município'].fillna(0) + df['% ISS Fora do Município'].fillna(0)

        razao_cliente = df['Empresa'].iloc[0]
        cnpj_cliente = df['Cnpj Empresa'].iloc[0]
        data_comp = pd.to_datetime(df['Data Competência'].iloc[0])
        comp_formatada = data_comp.strftime('%m.%Y')
        comp_titulo = data_comp.strftime('%m/%Y')

        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            m_base = {'Emissão NFe': 'Data Emissão', 'Número NFe': 'Nota Fiscal', 'Serviço Federal': 'Cód. Serviço', 'Prestador': 'Prestador', 'Cnpj/Cpf Prestador': 'CNPJ', 'Valor NFe': 'Vlr Contábil'}
            m_1708 = {**m_base, 'Base de Cálculo ISS': 'Base IRRF', '% IRRF': 'Aliq. IRRF', 'Valor IRRF': 'Valor IRRF'}
            m_csrf = {**m_base, 'Base de Cálculo ISS': 'Base CSR', '% CSRF': 'Aliq. CSRF', 'Valor CSRF': 'Total PCC'}
            m_8045 = {**m_base, 'Base de Cálculo ISS': 'Base IRRF', '% IRRF': 'Aliq. IRRF', 'Valor IRRF': 'Valor IRRF'}
            m_iss = {**m_base, 'BASE_ISS_TOTAL': 'Base ISS', 'ALIQ_ISS_TOTAL': 'Aliq. ISS', 'ISS_TOTAL': 'ISS'}
            m_inss = {**m_base, 'Base de Cálculo INSS': 'Base INSS', '% INSS': 'Aliq. INSS', 'Valor INSS': 'Valor INSS'}

            aplicar_estilo_ppc(writer, df[df['DARF IRRF'] == 1708], m_1708, 'IRRF 1708', 'IRRF 1708', razao_cliente, cnpj_cliente, comp_titulo)
            aplicar_estilo_ppc(writer, df[df['DARF CSRF'] == 5952], m_csrf, 'CSRF', 'CSRF', razao_cliente, cnpj_cliente, comp_titulo)
            aplicar_estilo_ppc(writer, df[df['DARF IRRF'] == 8045], m_8045, 'IRRF 8045', 'IRRF 8045', razao_cliente, cnpj_cliente, comp_titulo)
            aplicar_estilo_ppc(writer, df[df['ISS_TOTAL'] > 0], m_iss, 'ISS', 'ISS', razao_cliente, cnpj_cliente, comp_titulo)
            aplicar_estilo_ppc(writer, df[df['Valor INSS'] > 0], m_inss, 'INSS', 'INSS', razao_cliente, cnpj_cliente, comp_titulo)

        st.success(f"✅ Arquivo de **{razao_cliente}** processado com sucesso!")
        
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            nome_arquivo_saida = f"{razao_cliente} - Memoria de Calculo Retidos {comp_formatada}.xlsx"
            st.download_button(
                label="📥 Baixar Memória de Cálculo Formatada",
                data=output.getvalue(),
                file_name=nome_arquivo_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
