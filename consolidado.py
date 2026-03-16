import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(page_title="PPC - Gerador Duplo", page_icon="📊")

st.title("📊 Gerador: Memória + Consolidado")
st.markdown("Arraste a base do Unicont para gerar os dois relatórios.")

# --- SIDEBAR PARA ENTRADA DE DADOS MANUAIS ---
st.sidebar.header("Dados Adicionais (Relatório Consolidado)")
st.sidebar.markdown("Informe os valores que não estão no Unicont:")

inss_folha = st.sidebar.number_input("Valor INSS Folha", min_value=0.0, step=0.01)
irrf_trabalho = st.sidebar.number_input("IRRF (0561/0588)", min_value=0.0, step=0.01)
irrf_aluguel = st.sidebar.number_input("IRRF Aluguel (3208)", min_value=0.0, step=0.01)
responsavel = st.sidebar.text_input("Responsável pelo Relatório", "Marcos Paulo")

# Área de Upload
arquivo_upload = st.file_uploader("Selecione o arquivo Unicont (xlsx)", type=["xlsx"])

# --- FUNÇÃO ESTILO MEMÓRIA (OURO) ---
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

    if df_filtrado.empty:
        ws.merge_cells(start_row=7, start_column=2, end_row=7, end_column=10)
        cell_msg = ws.cell(row=7, column=2)
        cell_msg.value = "SEM MOVIMENTO"
        cell_msg.alignment = align_center
        for col_idx in range(2, 11): ws.cell(row=7, column=col_idx).border = thin_border
        return 0
    else:
        for col_orig, col_dest in colunas_mapeadas.items():
            if col_dest == 'Cód. Serviço':
                df_filtrado[col_orig] = df_filtrado[col_orig].astype(str).str.replace(',', '.')

        dados_finais = df_filtrado[list(colunas_mapeadas.keys())].rename(columns=colunas_mapeadas)
        moeda_cols = ['Vlr Contábil', 'Base IRRF', 'Valor IRRF', 'Base CSR', 'Total PCC', 'ISS', 'Valor INSS', 'Base INSS', 'Base ISS']
        
        for r_idx, row in enumerate(dados_finais.values, 7):
            for c_idx, value in enumerate(row, 2):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value
                cell.border = thin_border
                header_text = list(colunas_mapeadas.values())[c_idx-2]
                if header_text in moeda_cols: cell.number_format = 'R$ #,##0.00'
                elif 'Data' in header_text: cell.number_format = 'dd/mm/yyyy'
                elif 'Aliq' in header_text or '%' in header_text: cell.number_format = '0.00%'

        last_row = 6 + len(dados_finais)
        row_total = last_row + 1
        ws.cell(row=row_total, column=2, value="TOTAL").font = Font(bold=True)
        for col_idx in range(7, len(colunas_mapeadas) + 2):
            header_text = list(colunas_mapeadas.values())[col_idx-2]
            if header_text in moeda_cols:
                col_letter = get_column_letter(col_idx)
                ws.cell(row=row_total, column=col_idx, value=f"=SUM({col_letter}7:{col_letter}{last_row})").font = Font(bold=True)
                ws.cell(row=row_total, column=col_idx).number_format = 'R$ #,##0.00'
        
        # Retorna o total somado para o consolidado
        col_valor_final = 'Valor IRRF' if 'Valor IRRF' in colunas_mapeadas.values() else \
                         'Total PCC' if 'Total PCC' in colunas_mapeadas.values() else \
                         'ISS' if 'ISS' in colunas_mapeadas.values() else 'Valor INSS'
        return df_filtrado[list(colunas_mapeadas.keys())[list(colunas_mapeadas.values()).index(col_valor_final)]].sum()

# --- FUNÇÃO CONSOLIDADO ---
def gerar_consolidado(writer, dados_impostos, razao, cnpj, comp, resp):
    ws = writer.book.create_sheet("Consolidado Mensal")
    writer.sheets["Consolidado Mensal"] = ws
    ws.sheet_view.showGridLines = False
    
    # Estilos
    header_font = Font(bold=True, size=14)
    label_font = Font(bold=True)
    
    ws.cell(row=2, column=2, value="DCTFWeb - Relatório Mensal de Impostos Federais Consolidados").font = header_font
    ws.cell(row=4, column=2, value="Razão Social:").font = label_font
    ws.cell(row=4, column=4, value=razao)
    ws.cell(row=5, column=2, value="CNPJ:").font = label_font
    ws.cell(row=5, column=4, value=cnpj)
    ws.cell(row=6, column=2, value="Período de Apuração:").font = label_font
    ws.cell(row=6, column=4, value=comp)
    ws.cell(row=7, column=2, value="Responsável:").font = label_font
    ws.cell(row=7, column=4, value=resp)

    headers = ["Tipo", "Código Retenção", "Valor Retenção", "Descrição", "Observações"]
    for i, h in enumerate(headers, 2):
        cell = ws.cell(row=9, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color='002060', end_color='002060', fill_type='solid')

    row_idx = 10
    for imposto, valor in dados_impostos.items():
        ws.cell(row=row_idx, column=2, value=imposto.split('-')[0]) # Tipo
        ws.cell(row=row_idx, column=3, value=imposto.split('-')[1]) # Código
        cell_v = ws.cell(row=row_idx, column=4, value=valor)
        cell_v.number_format = 'R$ #,##0.00'
        ws.cell(row=row_idx, column=5, value=imposto.split('-')[2]) # Descrição
        ws.cell(row=row_idx, column=6, value="Memória de Cálculo" if valor > 0 else "Sem Movimento")
        row_idx += 1

# Execução Principal
if arquivo_upload:
    try:
        df = pd.read_excel(arquivo_upload)
        df['ISS_TOTAL'] = df['ISS Dentro do Município'].fillna(0) + df['ISS Fora do Município'].fillna(0)
        df['BASE_ISS_TOTAL'] = df['Base de Cálculo ISS'].fillna(0)
        df['ALIQ_ISS_TOTAL'] = df['% ISS Dentro do Município'].fillna(0) + df['% ISS Fora do Município'].fillna(0)

        razao = df['Empresa'].iloc[0]
        cnpj = df['Cnpj Empresa'].iloc[0]
        data_c = pd.to_datetime(df['Data Competência'].iloc[0])
        comp_f = data_c.strftime('%m.%Y')
        comp_t = data_c.strftime('%m/%Y')

        # 1. Processar Memória e Capturar Totais
        out_memoria = BytesIO()
        totais = {}
        with pd.ExcelWriter(out_memoria, engine='openpyxl') as writer:
            m_base = {'Emissão NFe': 'Data Emissão', 'Número NFe': 'Nota Fiscal', 'Serviço Federal': 'Cód. Serviço', 'Prestador': 'Prestador', 'Cnpj/Cpf Prestador': 'CNPJ', 'Valor NFe': 'Vlr Contábil'}
            
            totais['IRRF-1708-Serviços PJ'] = aplicar_estilo_ppc(writer, df[df['DARF IRRF'] == 1708], {**m_base, 'Base de Cálculo ISS': 'Base IRRF', 'Valor IRRF': 'Valor IRRF'}, 'IRRF 1708', 'IRRF 1708', razao, cnpj, comp_t)
            totais['CSRF-5952-Retenção PCC'] = aplicar_estilo_ppc(writer, df[df['DARF CSRF'] == 5952], {**m_base, 'Base de Cálculo ISS': 'Base CSR', 'Valor CSRF': 'Total PCC'}, 'CSRF', 'CSRF', razao, cnpj, comp_t)
            totais['IRRF-8045-Outros Rend.'] = aplicar_estilo_ppc(writer, df[df['DARF IRRF'] == 8045], {**m_base, 'Base de Cálculo ISS': 'Base IRRF', 'Valor IRRF': 'Valor IRRF'}, 'IRRF 8045', 'IRRF 8045', razao, cnpj, comp_t)
            totais['INSS-1162-Retenção NFSe'] = aplicar_estilo_ppc(writer, df[df['Valor INSS'] > 0], {**m_base, 'Base de Cálculo INSS': 'Base INSS', 'Valor INSS': 'Valor INSS'}, 'INSS', 'INSS', razao, cnpj, comp_t)
        
        # Adicionar manuais ao dicionário de totais
        totais['INSS-Folha-eSocial'] = inss_folha
        totais['IRRF-0561-Trabalho'] = irrf_trabalho
        totais['IRRF-3208-Aluguel'] = irrf_aluguel

        # 2. Gerar Consolidado
        out_consolidado = BytesIO()
        with pd.ExcelWriter(out_consolidado, engine='openpyxl') as writer:
            gerar_consolidado(writer, totais, razao, cnpj, comp_t, responsavel)

        st.success(f"✅ Documentos de {razao} prontos!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Memória de Cálculo", out_memoria.getvalue(), f"Memoria_{comp_f}.xlsx")
        with col2:
            st.download_button("📥 Relatório Consolidado", out_consolidado.getvalue(), f"Consolidado_{comp_f}.xlsx")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
