import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

def gerar_estrutura_relatorio():
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório Consolidado"
    
    # Configurações de Estilo
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                    top=Side(style='thin'), bottom=Side(style='thin'))
    fill_header = PatternFill(start_color='E6E6E6', end_color='E6E6E6', fill_type='solid')

    # Dados de exemplo para estruturação (Baseado no seu arquivo)
    #
    empresa = "REDCLOUD TECHNOLOGIES BRASIL SERVICOS DIGITAIS LTDA"
    cnpj = "47.597.633/0001-92"
    periodo = "Março de 2026"
    responsavel = "MARCOS PAULO SANTOS DE OLIVEIRA"

    # --- INÍCIO DA ESTRUTURA DO BLOCO ---
    
    # Título do Relatório
    ws.merge_cells('B16:G16')
    cell_titulo = ws['B16']
    cell_titulo.value = "DCTFWeb - Relatório Mensal de Impostos Federais Consolidados"
    cell_titulo.font = Font(bold=True, size=12)
    cell_titulo.alignment = center_align

    # Cabeçalho de Identificação
    campos = [
        ("Razão social", empresa),
        ("CNPJ", cnpj),
        ("Período de apuração", periodo),
        ("Responsável preenchimento", responsavel)
    ]

    for i, (label, valor) in enumerate(campos, 18):
        ws.cell(row=i, column=2, value=label).font = bold_font
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)
        ws.cell(row=i, column=5, value=valor).alignment = left_align

    # Cabeçalho da Tabela de Impostos[cite: 1]
    headers = ["Tipo", "Código Retenção RFB", "Valor Retenção", "Descrição do Código da Receita", "", "Observações"]
    for col, text in enumerate(headers, 2):
        cell = ws.cell(row=24, column=col, value=text)
        cell.font = bold_font
        cell.fill = fill_header
        cell.border = border
        cell.alignment = center_align

    # Lista de Impostos (Estrutura Fixa conforme arquivo)[cite: 1]
    impostos = [
        ("INSS", "Folha", "Informação transmitida via eSocial", "Considerar evidência enviada pelo RH"),
        ("IRRF", "0588", "IRRF - Rendimento do Trabalho sem Vínculo Empregatício", ""),
        ("IRRF", "0561", "IRRF - Rendimento do Trabalho Assalariado", ""),
        ("INSS", "1162", "Informação transmitida via EFD REINF - Retenção na fonte NFSe", "Considerar memória de cálculo do fiscal"),
        ("IRRF", "1708", "IRRF - Remuneração Serviços Prestados por Pessoa Jurídica", ""),
        ("IRRF", "8045", "IRRF - Outros Rendimentos", ""),
        ("IRRF", "3208", "IRRF - Aluguéis e Royalties Pagos a Pessoa Física", ""),
        ("IRRF", "3280", "IRRF - Rem Serv Prest Associad Coop Trabalho", ""),
        ("CSRF", "5952", "Retenção de Contribuições sobre Pagamentos de PJ a PJ", ""),
        ("IRRF", "0422", "IRRF - Royalties e Assistência Técnica - Exterior", ""),
        ("PIS", "8109", "PIS - FATURAMENTO - PJ EM GERAL", ""),
        ("COFINS", "2172", "COFINS - FATURAMENTO/PJ EM GERAL", ""),
        ("IRPJ", "2089", "IRPJ - LUCRO PRESUMIDO", ""),
        ("CSLL", "2372", "CSLL - LUCRO PRESUMIDO OU ARBITRADO", "")
    ]

    row_idx = 25
    for tipo, cod, desc, obs in impostos:
        ws.cell(row=row_idx, column=2, value=tipo).border = border
        ws.cell(row=row_idx, column=3, value=cod).border = border
        ws.cell(row=row_idx, column=4, value=0).border = border # Valor Inicial
        ws.cell(row=row_idx, column=5, value=desc).border = border
        ws.cell(row=row_idx, column=7, value=obs).border = border
        row_idx += 1

    # Totalizador[cite: 1]
    ws.cell(row=row_idx + 1, column=2, value="Valor Total DARF DCTFWeb").font = bold_font
    ws.cell(row=row_idx + 1, column=4, value=0).font = bold_font
    ws.cell(row=row_idx + 1, column=4).number_format = '"R$ "#,##0.00'

    # Ajuste de largura das colunas
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 60
    ws.column_dimensions['G'].width = 40

    return wb

# Streamlit Interface
st.title("Estruturador de Relatório Consolidado")

if st.button("Gerar Estrutura Base"):
    wb = gerar_estrutura_relatorio()
    output = BytesIO()
    wb.save(output)
    
    st.download_button(
        label="📥 Baixar Estrutura do Relatório",
        data=output.getvalue(),
        file_name="Estrutura_Relatorio_Consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
