import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(page_title="PPC - Sistema Fiscal", page_icon="📊", layout="wide")

# --- ESTILOS PARA A MEMÓRIA DE CÁLCULO ---
def aplicar_estilo_ppc(writer, df_filtrado, colunas_mapeadas, nome_aba, titulo_imposto, razao, cnpj, comp):
    ws = writer.book.create_sheet(nome_aba)
    writer.sheets[nome_aba] = ws
    ws.sheet_view.showGridLines = False 
    ws.column_dimensions['A'].width = 3
    fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
    font_branca = Font(color='FFFFFF', bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center')

    textos_cabecalho = [f'RAZÃO SOCIAL: {razao}', f'CNPJ: {cnpj}', f'{titulo_imposto} - COMPETÊNCIA {comp}']
    ultima_col_idx = len(colunas_mapeadas) + 1
    for row_num, texto in enumerate(textos_cabecalho, 2):
        ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=ultima_col_idx)
        cell = ws.cell(row=row_num, column=2, value=texto)
        cell.alignment = align_center
        cell.font = Font(bold=True, size=12)

    for col_num, header in enumerate(colunas_mapeadas.values(), 2):
        cell = ws.cell(row=6, column=col_num, value=header)
        cell.fill = fill_azul
        cell.font = font_branca
        cell.alignment = align_center
        cell.border = thin_border

    if df_filtrado.empty:
        ws.merge_cells(start_row=7, start_column=2, end_row=7, end_column=ultima_col_idx)
        cell_msg = ws.cell(row=7, column=2, value="SEM MOVIMENTO")
        cell_msg.font = Font(bold=True)
        cell_msg.alignment = align_center
    else:
        dados_finais = df_filtrado[list(colunas_mapeadas.keys())].rename(columns=colunas_mapeadas)
        for r_idx, row in enumerate(dados_finais.values, 7):
            for c_idx, value in enumerate(row, 2):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = thin_border
                cell.alignment = align_center

# --- FUNÇÃO 1: GERADOR DE MEMÓRIA ---
def aba_memoria_calculo():
    st.title("📊 Gerador de Memória de Cálculo")
    arquivo_upload = st.file_uploader("Arraste a planilha Unicont aqui", type=["xlsx"], key="memoria")
    if arquivo_upload:
        df = pd.read_excel(arquivo_upload)
        # Lógica de processamento simplificada para o teste
        st.success("Arquivo carregado com sucesso!")
        # (Aqui ficaria o restante da sua lógica da memória que já funciona)

# --- FUNÇÃO 2: RELATÓRIO CONSOLIDADO ---
def aba_relatorio_consolidado():
    st.title("📄 Relatório Mensal Consolidado")
    st.markdown("Estrutura idêntica ao modelo oficial da empresa.")
    
    if st.button("Gerar Estrutura de Teste"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Relatório Consolidado"
        ws.sheet_view.showGridLines = False

        # Estilos conforme sua solicitação
        cor_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
        cor_cinza = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        cor_laranja = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # Título[cite: 1]
        ws.column_dimensions['A'].width = 2
        ws.merge_cells('B2:G2')
        ws['B2'] = "DCTFWeb - Relatório Mensal de Impostos Federais Consolidados"
        ws['B2'].font = Font(color='FFFFFF', bold=True)
        ws['B2'].fill = cor_azul
        ws['B2'].alignment = Alignment(horizontal='center')

        # Cabeçalhos de Identificação Mesclados[cite: 1]
        ident = [("Razão social", "EMPRESA TESTE LTDA"), ("CNPJ", "00.000.000/0001-00"), ("Período", "Janeiro/2025")]
        for i, (lab, val) in enumerate(ident, 3):
            ws.merge_cells(f'B{i}:D{i}')
            ws[f'B{i}'] = lab
            ws[f'B{i}'].fill = cor_cinza
            ws.merge_cells(f'E{i}:G{i}')
            ws[f'E{i}'] = val

        # Tabela[cite: 1]
        headers = ["Tipo", "Código", "Valor", "Descrição", "", "Observações"]
        for col, text in enumerate(headers, 2):
            cell = ws.cell(row=8, column=col, value=text)
            cell.fill = cor_azul
            cell.font = Font(color='FFFFFF', bold=True)
            cell.border = border
        ws.merge_cells('E8:F8')

        # Linha de Exemplo (IRRF 1708)[cite: 1]
        ws.cell(row=9, column=2, value="IRRF").fill = cor_laranja
        ws.cell(row=9, column=3, value="1708")
        ws.merge_cells('E9:F9')
        ws.cell(row=9, column=5, value="IRRF - Remuneração Serviços PJ")
        for c in range(2, 8): ws.cell(row=9, column=c).border = border

        # Total[cite: 1]
        ws.merge_cells('B10:C10')
        ws['B10'] = "Valor Total DARF"
        ws['B10'].fill = cor_azul
        ws['B10'].font = Font(color='FFFFFF', bold=True)
        ws['D10'] = 0
        ws['D10'].font = Font(bold=True, color='FF0000')

        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Baixar Planilha", output.getvalue(), "Relatorio.xlsx")

# --- LÓGICA DE NAVEGAÇÃO (O que faz o site aparecer) ---
st.sidebar.title("Menu PPC")
escolha = st.sidebar.radio("Ir para:", ["Memória de Cálculo", "Relatório Consolidado"])

if escolha == "Memória de Cálculo":
    aba_memoria_calculo()
else:
    aba_relatorio_consolidado()
