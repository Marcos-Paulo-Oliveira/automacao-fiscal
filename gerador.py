import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO E ESTILOS ---
st.set_page_config(page_title="PPC - Automação Fiscal", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; background-color: #002060; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { width: 100%; background-color: #28a745; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE ESTILO (MEMÓRIA DE CÁLCULO) ---
def aplicar_estilo_ppc(writer, df_filtrado, colunas_mapeadas, nome_aba, titulo_imposto, razao, cnpj, comp):
    ws = writer.book.create_sheet(nome_aba)
    writer.sheets[nome_aba] = ws
    ws.sheet_view.showGridLines = False 
    ws.column_dimensions['A'].width = 3
    fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    font_branca = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    for row_num, texto in enumerate([f'RAZÃO SOCIAL: {razao}', f'CNPJ: {cnpj}', f'{titulo_imposto} - COMPETÊNCIA {comp}'], 2):
        ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=10)
        cell = ws.cell(row=row_num, column=2); cell.value = texto
        cell.alignment = align_center; cell.font = Font(bold=True, size=12)

    for col_num, header in enumerate(colunas_mapeadas.values(), 2):
        cell = ws.cell(row=6, column=col_num); cell.value = header
        cell.fill = fill_azul; cell.font = font_branca; cell.alignment = align_center; cell.border = thin_border

    if not df_filtrado.empty:
        dados = df_filtrado[list(colunas_mapeadas.keys())].rename(columns=colunas_mapeadas)
        for r_idx, row in enumerate(dados.values, 7):
            for c_idx, value in enumerate(row, 2):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value; cell.border = thin_border
                if any(x in str(list(colunas_mapeadas.values())[c_idx-2]) for x in ['Vlr', 'Base', 'Valor', 'Total', 'ISS']):
                    cell.number_format = 'R$ #,##0.00'

    for col in ws.columns:
        if col[0].column_letter != 'A':
            ws.column_dimensions[col[0].column_letter].width = 20

# --- FUNÇÃO NOVO RELATÓRIO CONSOLIDADO ---
def gerar_aba_consolidado(writer, df, razao, cnpj, comp_titulo, inputs_manuais):
    ws = writer.book.create_sheet("Consolidado Mensal")
    writer.sheets["Consolidado Mensal"] = ws
    ws.sheet_view.showGridLines = False
    
    fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    font_branca = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    info = [["Razão social", razao], ["CNPJ", cnpj], ["Período de apuração", comp_titulo]]
    for i, (label, val) in enumerate(info, 4):
        ws.cell(row=i, column=2, value=label).font = Font(bold=True)
        ws.cell(row=i, column=5, value=val)

    headers = ["Tipo", "Código Retenção RFB", "Valor Retenção", "Descrição do Código da Receita", "Observações"]
    for c, h in enumerate(headers, 2):
        cell = ws.cell(row=9, column=c, value=h)
        cell.fill = fill_azul; cell.font = font_branca; cell.border = thin_border; cell.alignment = Alignment(horizontal='center')

    soma_1708 = df[df['DARF IRRF'] == 1708]['Valor IRRF'].sum()
    soma_5952 = df[df['DARF CSRF'] == 5952]['Valor CSRF'].sum()
    soma_1162 = df['Valor INSS'].sum()

    linhas_relatorio = [
        ["INSS", "Folha", inputs_manuais.get('inss_folha', 0), "Informação via eSocial", "Evidência do RH"],
        ["IRRF", "0588", inputs_manuais.get('ir_0588', 0), "Rendimento Trabalho sem Vínculo", ""],
        ["IRRF", "0561", inputs_manuais.get('ir_0561', 0), "Rendimento Trabalho Assalariado", ""],
        ["INSS", "1162", soma_1162, "Retenção na fonte NFSe (EFD REINF)", "Memória de Cálculo Fiscal"],
        ["IRRF", "1708", soma_1708, "Remuneração Serviços Prestados PJ", ""],
        ["IRRF", "8045", inputs_manuais.get('ir_8045', 0), "Outros Rendimentos", ""],
        ["IRRF", "3208", inputs_manuais.get('ir_3208', 0), "Aluguéis e Royalties PF", ""],
        ["CSRF", "5952", soma_5952, "Retenção CSRF (PCC)", ""],
    ]

    for r, dados in enumerate(linhas_relatorio, 10):
        for c, valor in enumerate(dados, 2):
            cell = ws.cell(row=r, column=c, value=valor)
            cell.border = thin_border
            if c == 4: cell.number_format = 'R$ #,##0.00'

# --- INTERFACE ---
with st.sidebar:
    st.image("https://www.ppcaudit.com.br/wp-content/uploads/2017/07/logo-ppc.png", width=150)
    st.title("⚙️ Configurações")
    menu = st.radio("Selecione a ferramenta:", ["Memória de Cálculo", "Relatório Consolidado"])

st.title(f"📊 {menu}")

arquivo = st.file_uploader("Suba o arquivo UneCont (xlsx)", type="xlsx")

if arquivo:
    df = pd.read_excel(arquivo)
    razao = df['Empresa'].iloc[0]
    comp_titulo = pd.to_datetime(df['Data Competência'].iloc[0]).strftime('%m/%Y')
    
    inputs_manuais = {}
    if menu == "Relatório Consolidado":
        st.subheader("📝 Informações Complementares")
        st.info("Preencha os valores abaixo que não constam no relatório automático.")
        col1, col2 = st.columns(2)
        with col1:
            inputs_manuais['inss_folha'] = st.number_input("INSS Folha (eSocial)", value=0.0)
            inputs_manuais['ir_0588'] = st.number_input("IRRF 0588", value=0.0)
            inputs_manuais['ir_0561'] = st.number_input("IRRF 0561", value=0.0)
        with col2:
            inputs_manuais['ir_8045'] = st.number_input("IRRF 8045", value=0.0)
            inputs_manuais['ir_3208'] = st.number_input("IRRF 3208", value=0.0)

    if st.button("🚀 Gerar Documentos"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            m_base = {'Emissão NFe': 'Data Emissão', 'Número NFe': 'Nota Fiscal', 'Prestador': 'Prestador', 'Valor NFe': 'Vlr Contábil'}
            m_1708 = {**m_base, 'Valor IRRF': 'Valor IRRF'}
            m_csrf = {**m_base, 'Valor CSRF': 'Total PCC'}
            
            aplicar_estilo_ppc(writer, df[df['DARF IRRF'] == 1708], m_1708, 'Detalhamento 1708', 'IRRF 1708', razao, df['Cnpj Empresa'].iloc[0], comp_titulo)
            aplicar_estilo_ppc(writer, df[df['DARF CSRF'] == 5952], m_csrf, 'Detalhamento CSRF', 'CSRF', razao, df['Cnpj Empresa'].iloc[0], comp_titulo)
            gerar_aba_consolidado(writer, df, razao, df['Cnpj Empresa'].iloc[0], comp_titulo, inputs_manuais)

        st.success("Tudo pronto!")
        st.download_button("📥 Baixar Relatórios", data=output.getvalue(), file_name=f"Relatorio_Fiscal_{razao}.xlsx")
