import xlwings as xw
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def atualizar_excel():
    caminho_arquivo = r"Z:\\Controladoria\\Reservas\\Reservas.xlsx"
    wb = xw.Book(caminho_arquivo)
    wb.api.RefreshAll()
    wb.app.api.CalculateFullRebuild()
    wb.save()
    wb.close()
    print("Planilha atualizada e salva com sucesso!")
    return caminho_arquivo

def importar_para_sheets(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    df = df.fillna("")

    # 👉 Debug: mostra os nomes das colunas
    print("Colunas encontradas no Excel:", df.columns)

    # 👉 Cria coluna Nome Completo na ordem NOME + SOBRENOME
    if "NOME" in df.columns and "SOBRENOME" in df.columns:
        df["Nome Completo"] = df["NOME"].astype(str) + " " + df["SOBRENOME"].astype(str)

    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        r"Z:\\Controladoria\\Reservas\\solar-nimbus-477516-b0-80bc7621dd62.json", scope
    )
    client = gspread.authorize(creds)

    sheet = client.open("PYTHON_TESTE").worksheet("RESERVAS")

    values = [df.columns.values.tolist()] + df.astype(str).values.tolist()
    sheet.clear()
    sheet.update(values)

    print("Dados importados para o Google Sheets com sucesso!")

if __name__ == "__main__":
    caminho = atualizar_excel()
    importar_para_sheets(caminho)
