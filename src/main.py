import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from utils import get_email_gestor, enviar_email, PLANILHA_PATH

print(f"📂 Usando planilha: {PLANILHA_PATH}")

wb = load_workbook(PLANILHA_PATH)
ws = wb["2024"]

df_contratos = pd.read_excel(PLANILHA_PATH, sheet_name="2024", header=4)

fill_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

hoje = datetime.today()

for idx, row in df_contratos.iterrows():
    fim_vigencia = row["FIM DA VIGÊNCIA [14]"]
    fiscal = row["NOMES DO FISCAL CONTRATO [20]"]

    if pd.notnull(fim_vigencia):
        fim_vigencia = pd.to_datetime(fim_vigencia, errors="coerce")

        if pd.notnull(fim_vigencia):
            dias_restantes = (fim_vigencia - hoje).days

            # ✅ Ignora vencidos
            if 0 <= dias_restantes <= 90:
                # Pintar a linha no Excel
                excel_row = idx + 6  # dados começam na linha 6
                for cell in ws[excel_row]:
                    cell.fill = fill_red

                # Buscar email do gestor
                email = get_email_gestor(fiscal)
                if email:
                    mensagem = f"""
                    Prezado {fiscal},

                    O contrato "{row['Nº DO CONTRATO [10]']}" está a {dias_restantes} dias do fim da vigência ({fim_vigencia.date()}).

                    Favor tomar as devidas providências.
                    """
                    enviar_email(
                        email, "Alerta: Contrato próximo ao vencimento", mensagem
                    )

# Salvar planilha atualizada
wb.save(PLANILHA_PATH)
print("✅ Processo concluído com sucesso 🚀")
