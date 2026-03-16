import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Configuração da página
st.set_page_config(page_title="PPC - Consolidado Fiscal", page_icon="📝", layout="wide")

st.title("📝 Gerador de Relatório Consolidado")
st.info("Arraste sua **Memória de Cálculo (Base Ouro)** abaixo.")

arquivo_upload = st.file_uploader("Selecione a Memória de Cálculo (xlsx)", type=["xlsx"])

if arquivo_upload:
    try:
        xl = pd.ExcelFile(arquivo_upload)
        
        # Identificação da Empresa
        primeira_aba = xl.sheet_names[0]
        df_busca = xl.parse(primeira_aba, header=None).astype(str)
        razao_social, cnpj, competencia = "Não Encontrado", "Não Encontrado", "00/0000"

        for i in range(min(len(df_busca), 15)):
            for j in range(len(df_busca.columns)):
                celula = df_busca.iloc[i, j]
                if "RAZÃO SOCIAL:" in celula: razao_social = celula.replace("RAZÃO SOCIAL:", "").strip()
                elif "CNPJ:" in celula: cnpj = celula.replace("CNPJ:", "").strip()
                elif "COMPETÊNCIA" in celula: competencia = celula.split("COMPETÊNCIA")[-1].strip()

        st.success(f"📌 Empresa: **{razao_social}**")

        # Inputs Manuais
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: inss_folha = st.number_input("INSS Folha (eSocial)", min_value=0.0, step=0.01, format="%.2f")
        with col_m2: ir_0588 = st.number_input("IRRF 0588", min_value=0.0, step=0.01, format="%.2f")
        with col_m3: ir_0561 = st.number_input("IRRF 0561", min_value=0.0, step=0.01, format="%.2f")

        # FUNÇÃO PARA PEGAR APENAS A CÉLULA DO TOTAL
        def capturar_valor_total(nome_aba):
            # Normalização de nomes (ajuste para 'IRRF 1708 ' com espaço no final se houver)
            abas_reais = {n.strip(): n for n in xl.sheet_names}
            nome_real = abas_reais.get(nome_aba.strip())
            
            if nome_real:
                df = xl.parse(nome_real, header=None)
                # Procura a linha que contém a palavra "TOTAL"
                for idx, row in df.iterrows():
                    row_str = row.astype(str).values
                    if any("TOTAL" in s.upper() for s in row_str):
                        # Pega o último valor numérico dessa linha
                        valores_linha = pd.to_numeric(row, errors='coerce')
                        valor_final = valores_linha.dropna().iloc[-1] if not valores_linha.dropna().empty else 0.0
                        return float(valor_final)
            return 0.0

        # Mapeamento dos valores das abas
        valores = {
            "1708": capturar_valor_total("IRRF 1708"),
            "8045": capturar_valor_total("IRRF 8045"),
            "5952": capturar_valor_total("CSRF"),
            "1162": capturar_valor_total("INSS"),
            "3208": capturar_valor_total("IRRF 3208")
        }

        if st.button("🚀 Gerar Relatório Consolidado"):
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "Consolidado Mensal"
            ws.sheet_view.showGridLines = False

            # Estilos
            fill_azul = PatternFill(start_color='002060', end_color='002060', fill_type='solid')
            fill_cinza = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
            font_branca = Font(color='FFFFFF', bold=True)
            font_bold = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Cabeçalho
            ws.cell(row=2, column=2, value="DCTFWeb - Relatório Mensal de Impostos Federais Consolidados").font = Font(bold=True, size=14)
            info = [["Razão social", razao_social], ["CNPJ", cnpj], ["Período de apuração", competencia], ["Responsável", "Marcos Paulo Santos de Oliveira"]]
            for i, (label, val) in enumerate(info, 4):
                ws.cell(row=i, column=2, value=label).font = font_bold
                ws.cell(row=i, column=5, value=val)

            # Tabela
            headers = ["Tipo", "Código Retenção RFB", "Valor Retenção", "Descrição do Código da Receita", "Observações"]
            for c, h in enumerate(headers, 2):
                cell = ws.cell(row=9, column=c, value=h)
                cell.fill = fill_azul; cell.font = font_branca; cell.border = thin_border; cell.alignment = Alignment(horizontal='center')

            dados_finais = [
                ["INSS", "Folha", inss_folha, "Informação transmitida via eSocial", "Considerar evidência enviada pelo RH"],
                ["IRRF", "0588", ir_0588, "IRRF - Rendimento do Trabalho sem Vínculo Empregatício", ""],
                ["IRRF", "0561", ir_0561, "IRRF - Rendimento do Trabalho Assalariado", ""],
                ["INSS", "1162", valores["1162"], "Informação transmitida via EFD REINF - Retenção na fonte NFSe", "Considerar memória de cálculo do fiscal"],
                ["IRRF", "1708", valores["1708"], "IRRF - Remuneração Serviços Prestados por Pessoa Jurídica", ""],
                ["IRRF", "8045", valores["8045"], "IRRF - Outros Rendimentos", ""],
                ["IRRF", "3208", valores["3208"], "IRRF - Aluguéis e Royalties Pagos a Pessoa Física", ""],
                ["CSRF", "5952", valores["5952"], "Retenção de Contribuições (PIS/COFINS/CSLL) - Retenção CSRF (PCC)", ""]
            ]

            row_idx = 10
            for linha in dados_finais:
                for c_idx, valor in enumerate(linha, 2):
                    cell = ws.cell(row=row_idx, column=c_idx, value=valor)
                    cell.border = thin_border
                    if c_idx == 4: cell.number_format = 'R$ #,##0.00'
                row_idx += 1

            # --- LINHA DO TOTAL GERAL (DARF WEB) ---
            total_geral = inss_folha + ir_0588 + ir_0561 + sum(valores.values())
            
            ws.cell(row=row_idx, column=2, value="TOTAL").font = font_bold
            ws.cell(row=row_idx, column=2).border = thin_border
            ws.cell(row=row_idx, column=3, value="DARF WEB").font = font_bold
            ws.cell(row=row_idx, column=3).border = thin_border
            
            cell_total = ws.cell(row=row_idx, column=4, value=total_geral)
            cell_total.font = font_bold
            cell_total.fill = fill_cinza
            cell_total.border = thin_border
            cell_total.number_format = 'R$ #,##0.00'
            
            # Preencher o resto da linha do total com bordas
            ws.cell(row=row_idx, column=5).border = thin_border
            ws.cell(row=row_idx, column=6).border = thin_border

            # Ajustes de coluna
            larguras = [12, 18, 20, 55, 35]
            for i, largura in enumerate(larguras, 2):
                ws.column_dimensions[get_column_letter(i)].width = largura

            wb.save(output)
            st.download_button(label="📥 Baixar Relatório Consolidado", data=output.getvalue(), 
                               file_name=f"Consolidado - {razao_social}.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro: {e}")
