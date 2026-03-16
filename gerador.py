import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
from fpdf import FPDF
import zipfile

# Configuração da página
st.set_page_config(page_title="PPC - Gerador", page_icon="📊", layout="wide")

st.title("📊 Gerador de Memória de Cálculo")
st.markdown("Arraste a planilha exportada do sistema abaixo:")

arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

# --- FUNÇÃO PDF CORRIGIDA (COM AUTO-AJUSTE E FIT TO PAGE) ---
def criar_pdf_aba(df, titulo, razao, cnpj, competencia):
    # Usamos Paisagem (L) para caber mais colunas
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Helvetica", 'B', 12)
    
    # Cabeçalho
    pdf.cell(0, 7, f"RAZÃO SOCIAL: {razao}", ln=True)
    pdf.cell(0, 7, f"CNPJ: {cnpj}", ln=True)
    pdf.cell(0, 7, f"{titulo} - COMPETÊNCIA {competencia}", ln=True)
    pdf.ln(5)
    
    if df.empty:
        pdf.set_font("Helvetica", 'B', 14)
        pdf.cell(0, 10, "SEM MOVIMENTO", border=1, ln=True, align='C')
    else:
        # 1. Calcular a largura necessária para cada coluna (Ajuste automático)
        pdf.set_font("Helvetica", '', 8)
        larguras_colunas = []
        for col in df.columns:
            # Mede o tamanho do header
            maior_texto = pdf.get_string_width(str(col)) + 6
            # Mede o tamanho dos dados da coluna
            for item in df[col]:
                tam_item = pdf.get_string_width(str(item)) + 6
                if tam_item > maior_texto:
                    maior_texto = tam_item
            larguras_colunas.append(maior_texto)
        
        # 2. Ajustar tudo em uma única página (Fit to page)
        largura_total_tabela = sum(larguras_colunas)
        largura_disponivel_a4 = 275 # Margens inclusas
        
        # Se a tabela for maior que a página, reduzimos proporcionalmente
        if largura_total_tabela > largura_disponivel_a4:
            fator_escala = largura_disponivel_a4 / largura_total_tabela
            larguras_colunas = [w * fator_escala for w in larguras_colunas]
        
        # 3. Desenhar o Cabeçalho da Tabela
        pdf.set_font("Helvetica", 'B', 8)
        pdf.set_fill_color(0, 32, 96) # Azul Marinho
        pdf.set_text_color(255, 255, 255) # Branco
        
        for i, col in enumerate(df.columns):
            pdf.cell(larguras_colunas[i], 7, str(col), border=1, align='C', fill=True)
        pdf.ln()
        
        # 4. Desenhar as Linhas
        pdf.set_font("Helvetica", '', 7)
        pdf.set_text_color(0, 0, 0)
        
        for _, row in df.iterrows():
            for i, val in enumerate(row):
                texto = str(val)
                # Formata moeda se necessário
                if "Valor" in df.columns[i] or "Total" in df.columns[i] or "Base" in df.columns[i]:
                    try:
                        texto = f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    except: pass
                
                pdf.cell(larguras_colunas[i], 6, texto, border=1, align='C')
            pdf.ln()
            
    return pdf.output()

# --- FUNÇÃO EXCEL (MANTIDA) ---
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
        ws.cell(row=row_total, column=2).alignment = Alignment(horizontal='right')
        for col_idx in range(7, len(colunas_mapeadas) + 2):
            header_text = list(colunas_mapeadas.values())[col_idx-2]
            if header_text in moeda_cols:
                col_letter = get_column_letter(col_idx)
                ws.cell(row=row_total, column=col_idx, value=f"=SUM({col_letter}7:{col_letter}{last_row})").font = Font(bold=True)
                ws.cell(row=row_total, column=col_idx).number_format = 'R$ #,##0.00'

    for col in ws.columns:
        column_letter = col[0].column_letter
        if column_letter != 'A': ws.column_dimensions[column_letter].width = 20

# Execução
if arquivo_upload:
    try:
        df = pd.read_excel(arquivo_upload)
        df['ISS_TOTAL'] = df['ISS Dentro do Município'].fillna(0) + df['ISS Fora do Município'].fillna(0)
        df['BASE_ISS_TOTAL'] = df['Base de Cálculo ISS'].fillna(0)
        df['ALIQ_ISS_TOTAL'] = df['% ISS Dentro do Município'].fillna(0) + df['% ISS Fora do Município'].fillna(0)

        razao_cliente = df['Empresa'].iloc[0]
        cnpj_cliente = df['Cnpj Empresa'].iloc[0]
        data_comp = pd.to_datetime(df['Data Competência'].iloc[0])
        comp_formatada = data_comp.strftime('%m.%Y')
        comp_titulo = data_comp.strftime('%m/%Y')

        # Mapeamentos
        m_base = {'Emissão NFe': 'Data Emissão', 'Número NFe': 'Nota Fiscal', 'Serviço Federal': 'Cód. Serviço', 'Prestador': 'Prestador', 'Cnpj/Cpf Prestador': 'CNPJ', 'Valor NFe': 'Vlr Contábil'}
        mapeamentos = {
            'IRRF 1708': (df[df['DARF IRRF'] == 1708], {**m_base, 'Base de Cálculo ISS': 'Base IRRF', 'Valor IRRF': 'Valor IRRF'}),
            'CSRF': (df[df['DARF CSRF'] == 5952], {**m_base, 'Base de Cálculo ISS': 'Base CSR', 'Valor CSRF': 'Total PCC'}),
            'IRRF 8045': (df[df['DARF IRRF'] == 8045], {**m_base, 'Base de Cálculo ISS': 'Base IRRF', 'Valor IRRF': 'Valor IRRF'}),
            'ISS': (df[df['ISS_TOTAL'] > 0], {**m_base, 'BASE_ISS_TOTAL': 'Base ISS', 'ISS_TOTAL': 'ISS'}),
            'INSS': (df[df['Valor INSS'] > 0], {**m_base, 'Base de Cálculo INSS': 'Base INSS', 'Valor INSS': 'Valor INSS'})
        }

        # 1. Excel
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for nome, (dados, mapeamento) in mapeamentos.items():
                aplicar_estilo_ppc(writer, dados, mapeamento, nome, nome, razao_cliente, cnpj_cliente, comp_titulo)
        
        st.success(f"✅ Planilha de {razao_cliente} pronta!")
        st.download_button(label="📥 Baixar Memória em Excel", data=output_excel.getvalue(), file_name=f"Memoria_{razao_cliente}_{comp_formatada}.xlsx")

        st.divider()

        # 2. PDF Confirmado
        st.subheader("📄 Exportação para Cliente (PDF)")
        if st.button("Confirmar dados e preparar PDFs para envio"):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for nome, (dados, mapeamento) in mapeamentos.items():
                    # Preparação correta dos dados para o PDF
                    if dados.empty:
                        df_pdf = pd.DataFrame()
                    else:
                        df_pdf = dados[list(mapeamento.keys())].rename(columns=mapeamento)
                    
                    pdf_content = criar_pdf_aba(df_pdf, nome, razao_cliente, cnpj_cliente, comp_titulo)
                    zip_file.writestr(f"{nome}.pdf", pdf_content)
            
            st.success("PDFs gerados com sucesso!")
            st.download_button(label="📥 Baixar Pack de PDFs (ZIP)", data=zip_buffer.getvalue(), file_name=f"PDFs_{razao_cliente}.zip")

    except Exception as e:
        st.error(f"Erro: {e}")
