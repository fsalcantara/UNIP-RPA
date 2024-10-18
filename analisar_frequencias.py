import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib.pyplot as plt
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import base64

# Configurar o acesso à planilha do Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
client = gspread.authorize(creds)

# Caminho para o arquivo JSON da conta de serviço
SERVICE_ACCOUNT_FILE = 'email-tcc-unip-58bcaef6d9aa.json'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
USER_TO_SEND_AS = 'uniprpa@gmail.com'

# Criar a pasta de gráficos, se não existir
pasta_graficos = "analise_grafica"
if not os.path.exists(pasta_graficos):
    os.makedirs(pasta_graficos)
    print(f"Pasta {pasta_graficos} criada.")

# Função para carregar a sheet e transformar em DataFrame
def carregar_frequencia(mes):
    sheet = client.open("planilha_funcionarios").worksheet(mes)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

# Função para calcular a análise da frequência
def analisar_frequencia(df):
    df['Presenças'] = df.iloc[:, 2:].apply(lambda row: (row == 'P').sum(), axis=1)
    df['Ausências'] = df.iloc[:, 2:].apply(lambda row: (row == 'A').sum(), axis=1)
    df['Justificativas'] = df.iloc[:, 2:].apply(lambda row: (row == 'J').sum(), axis=1)

    num_dias = df.shape[1] - 2
    df['Taxa de Presença (%)'] = (df['Presenças'] / num_dias) * 100
    df['Taxa de Ausência (%)'] = (df['Ausências'] / num_dias) * 100

    return df[['Nome do Funcionário', 'Área', 'Presenças', 'Ausências', 'Justificativas', 'Taxa de Presença (%)',
               'Taxa de Ausência (%)']]

# Função para gerar gráficos e salvar na pasta analise_grafica
def gerar_graficos(analise_df, mes):
    try:
        plt.figure(figsize=(10, 6))
        analise_df[['Nome do Funcionário', 'Presenças', 'Ausências']].set_index('Nome do Funcionário').plot(kind='bar', stacked=True)
        plt.title(f'Comparação de Presenças e Ausências - {mes}')
        plt.xlabel('Funcionários')
        plt.ylabel('Número de Dias')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        grafico_barras_path = os.path.join(pasta_graficos, f'grafico_presenca_ausencia_{mes}.png')
        plt.savefig(grafico_barras_path)
        plt.close()
        print(f"Gráfico de barras salvo em {grafico_barras_path}")

        plt.figure(figsize=(6, 6))
        totais = analise_df[['Presenças', 'Ausências', 'Justificativas']].sum()
        plt.pie(totais, labels=['Presenças', 'Ausências', 'Justificativas'], autopct='%1.1f%%', startangle=90,
                colors=['#4CAF50', '#FF5722', '#FFC107'])
        plt.title(f'Distribuição de Presenças, Ausências e Justificativas - {mes}')
        plt.axis('equal')

        grafico_pizza_path = os.path.join(pasta_graficos, f'grafico_distribuicao_{mes}.png')
        plt.savefig(grafico_pizza_path)
        plt.close()
        print(f"Gráfico de pizza salvo em {grafico_pizza_path}")

        return grafico_barras_path, grafico_pizza_path
    except Exception as e:
        print(f"Erro ao gerar gráficos: {e}")
        return None, None

# Função para gerar o relatório em texto
def gerar_relatorio(analise_df, mes):
    relatorio = f"Relatório de Frequência - {mes}\n"
    relatorio += "=" * 50 + "\n"
    for _, row in analise_df.iterrows():
        relatorio += (f"Funcionário: {row['Nome do Funcionário']} | Área: {row['Área']} | "
                      f"Presenças: {row['Presenças']} | Ausências: {row['Ausências']} | "
                      f"Justificativas: {row['Justificativas']} | Taxa de Presença: {row['Taxa de Presença (%)']:.2f}% | "
                      f"Taxa de Ausência: {row['Taxa de Ausência (%)']:.2f}%\n")
    relatorio += "=" * 50 + "\n"
    return relatorio

# Função para autenticar e obter o serviço Gmail API
def get_gmail_service():
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    delegated_credentials = credentials.with_subject(USER_TO_SEND_AS)
    service = build('gmail', 'v1', credentials=delegated_credentials)
    return service

# Função para criar a mensagem de e-mail
def create_message(to, subject, message_text):
    message = MIMEText(message_text)
    message['to'] = to
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

# Função para enviar e-mail usando a Gmail API
def enviar_email_via_gmail(service, destinatario, relatorio):
    try:
        message = create_message(destinatario, "Relatório de Frequência Mensal com Gráficos", relatorio)
        message = service.users().messages().send(userId=USER_TO_SEND_AS, body=message).execute()
        print(f"E-mail enviado com sucesso para {destinatario}, ID da mensagem: {message['id']}")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Função para criar ou atualizar a aba com a análise da frequência
def exportar_analise_para_sheets(sheet, nome_do_mes, analise_df):
    nome_da_aba = f"Análise {nome_do_mes}"

    try:
        aba_analise = sheet.worksheet(nome_da_aba)
        aba_analise.clear()
    except gspread.exceptions.WorksheetNotFound:
        aba_analise = sheet.add_worksheet(title=nome_da_aba, rows=len(analise_df) + 1, cols=len(analise_df.columns) + 1)

    aba_analise.update([analise_df.columns.values.tolist()] + analise_df.values.tolist())

# Função para realizar a análise de todos os meses
def analisar_todos_os_meses():
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
             'Novembro', 'Dezembro']

    sheet = client.open("planilha_funcionarios")

    service = get_gmail_service()  # Obter o serviço da API Gmail

    for mes in meses:
        print(f"Analisando {mes}...")
        df = carregar_frequencia(mes)
        analise_df = analisar_frequencia(df)

        exportar_analise_para_sheets(sheet, mes, analise_df)

        relatorio = gerar_relatorio(analise_df, mes)

        # Gerar gráficos
        grafico_barras_path, grafico_pizza_path = gerar_graficos(analise_df, mes)

        # Enviar e-mail usando a Gmail API
        destinatario = input()
        enviar_email_via_gmail(service, destinatario, relatorio)

        print(f"Análise de {mes} exportada e relatório com gráficos enviado com sucesso!\n")

# Executar a análise para todos os meses
analisar_todos_os_meses()
