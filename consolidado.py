import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

def gerador_relatorio_consolidado():
    st.title("📄 Relatório Mensal Consolidado")
    st.markdown("Clique no botão abaixo para gerar a estrutura idêntica ao modelo da empresa.")
    
    if st.button("Gerar Estrutura Corrigida"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Relatório Consolidado"
        ws.sheet_view.showGridLines = False

        # --- CONFIGURAÇÃO DE CORES E ESTILOS ---
        cor_azul_cabecalho = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
        cor_cinza_claro = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        cor_laranja_tipo = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
        font_branca_bold = Font(color='FFFFFF', bold=True)
        font_preta_bold = Font(color='000000', bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Ajuste de largura inicial
        ws.column_dimensions['A'].width = 2
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 60
        ws.column_dimensions['G'].width = 40

        # --- 1. TÍTULO PRINCIPAL (Linha 2) ---
        ws.merge_cells('B2:G2')
        cell_titulo = ws['B2']
        cell_titulo.value = "DCTFWeb - Relatório Mensal de Impostos Federais Consolidados"
        cell_titulo.font = font_branca_bold
        cell_titulo.fill = cor_azul_cabecalho
        cell_titulo.alignment = Alignment(horizontal='center', vertical='center')

        # --- 2. IDENTIFICAÇÃO (Linhas 3 a 6) ---
        identificacao = [
            ("Razão social", "REDCLOUD TECHNOLOGIES BRASIL SERVICOS DIGITAIS LTDA"),
            ("CNPJ", "47.597.633/0001-92"),
            ("Período de apuração", "Fevereiro de 2026"),
            ("Responsável preenchimento", "MARCOS PAULO SANTOS DE OLIVEIRA")
        ]

        for i, (label, valor) in enumerate(identificacao, 3):
            # Mescla as colunas B, C e D para o rótulo
            ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
            cell_label = ws.cell(row=i, column=2, value=label)
            cell_label.font = font_preta_bold
            cell_label.fill = cor_cinza_claro
            cell_label.alignment = Alignment(horizontal='left')

            # Mescla as colunas E, F e G para o valor
            ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=7)
            cell_valor = ws.cell(row=i, column=5, value=valor)
            cell_valor.alignment = Alignment(horizontal='left')

        # --- 3. CABEÇALHO DA TABELA (Linha 9) ---
        headers = ["Tipo", "Código Retenção", "Valor Retenção", "Descrição do Código da Receita", "", "Observações"]
        # Nota: Ajustei as posições para bater com o print (B, C, D, E/F mesclados, G)
        cols_pos = [2, 3, 4, 5, 6, 7]
        for col, text in zip(cols_pos, headers):
            cell = ws.cell(row=9, column=col, value=text)
            cell.font = font_branca_bold
            cell.fill = cor_azul_cabecalho
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        # Mescla a descrição (Colunas E e F)
        ws.merge_cells('E9:F9')

        # --- 4. LISTA DE IMPOSTOS ---
        impostos = [
            ("INSS", "Folha", "Informação transmitida via eSocial", ""),
            ("IRRF", "0588", "IRRF - Rendimento do Trabalho sem Vínculo Empregatício", ""),
            ("IRRF", "0561", "IRRF - Rendimento do Trabalho Assalariado", ""),
            ("INSS", "1162", "Informação transmitida via EFD REINF - Retenção na fonte NFSe", ""),
            ("IRRF", "1708", "IRRF - Remuneração Serviços Prestados por Pessoa Jurídica", ""),
            ("IRRF", "8045", "IRRF - Outros Rendimentos", ""),
            ("IRRF", "3208", "IRRF - Aluguéis e Royalties Pagos a Pessoa Física", ""),
            ("IRRF", "3280", "IRRF - Rem Serv Prest Associad Coop Trabalho", ""),
            ("CSRF", "5952", "Retenção de Contribuições (CSLL, Cofins e PIS)", ""),
            ("IRRF", "0422", "IRRF - Royalties e Assistência Técnica - Exterior", ""),
            ("PIS", "8109", "PIS - FATURAMENTO - PJ EM GERAL", ""),
            ("COFINS", "2172", "COFINS - FATURAMENTO/PJ EM GERAL", ""),
            ("IRPJ", "2089", "IRPJ - LUCRO PRESUMIDO", ""),
            ("CSLL", "2372", "CSLL - LUCRO PRESUMIDO OU ARBITRADO", "")
        ]

        row_idx = 10
        for tipo, cod, desc, obs in impostos:
            # Coluna Tipo (Laranja)
            c_tipo = ws.cell(row=row_idx, column=2, value=tipo)
            c_tipo.fill = cor_laranja_tipo
            c_tipo.font = font_preta_bold
            
            # Outras Colunas
            ws.cell(row=row_idx, column=3, value=cod)
            ws.cell(row=row_idx, column=4, value=0).number_format = 'R$ #,##0.00'
            
            # Descrição mesclada E e F
            ws.merge_cells(start_row=row_idx, start_column=5, end_row=row_idx, end_column=6)
            ws.cell(row=row_idx, column=5, value=desc)
            
            ws.cell(row=row_idx, column=7, value=obs)

            # Aplicar bordas e centralização em tudo
            for c in range(2, 8):
                cell = ws.cell(row=row_idx, column=c)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            row_idx += 1

        # --- 5. VALOR TOTAL (Logo abaixo da última linha) ---
        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=3)
        cell_total_label = ws.cell(row=row_idx, column=2, value="Valor Total DARF")
        cell_total_label.font = font_branca_bold
        cell_total_label.fill = cor_azul_cabecalho
        cell_total_label.alignment = Alignment(horizontal='center')
        cell_total_label.border = thin_border

        cell_total_valor = ws.cell(row=row_idx, column=4, value=0)
        cell_total_valor.font = Font(bold=True, color='FF0000') # Vermelho igual ao print
        cell_total_valor.number_format = 'R$ #,##0.00'
        cell_total_valor.border = thin_border
        cell_total_valor.alignment = Alignment(horizontal='center')

        # Download
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Estrutura Consolidada",
            data=output.getvalue(),
            file_name="Relatorio_Consolidado_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
