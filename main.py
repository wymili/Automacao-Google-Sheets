import os
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

sys.stdout.reconfigure(encoding='utf-8')

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1gds10tb4bSCEoaEujCO64UQDycg1wmWoNokH4_Rge-4"

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        # Obtém e imprime o título das abas para confirmar o nome correto
        spreadsheet_info = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_titles = [sheet['properties']['title'] for sheet in spreadsheet_info['sheets']]
        print("Abas disponíveis:", sheet_titles)

        # Confirme o nome exato da aba
        sheet_name = sheet_titles[0]  # Ou ajuste para a aba específica se houver mais de uma
        print(f"Tentando acessar a aba: {sheet_name}")

        for row in range(2, 8):
            num1 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!A{row}").execute().get("values")[0][0])
            num2 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!B{row}").execute().get("values")[0][0])
            calculation_result = num1 + num2
            print(f"Processing {num1} + {num2}")

            # Atualiza a célula com o resultado do cálculo
            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!C{row}",valueInputOption="USER_ENTERED",body={"values": [[str(calculation_result)]]}).execute()
            print(calculation_result)
            # Marca a operação como "Done" na coluna D
            sheets.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{sheet_name}!D{row}",
                valueInputOption="USER_ENTERED",
                body={"values": [["Done"]]}
            ).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == "__main__":
    main()
