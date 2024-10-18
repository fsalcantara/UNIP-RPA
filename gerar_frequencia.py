import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import random

# Configurar o acesso à planilha do Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json',
                                                         scope)
client = gspread.authorize(creds)

# Abrir a planilha pelo nome (substitua pelo nome da sua planilha)
sheet = client.open("planilha_funcionarios")

# Ler a planilha de funcionários para um DataFrame
worksheet_funcionarios = sheet.worksheet("Planilha funcionários")
data = worksheet_funcionarios.get_all_records()
funcionarios_df = pd.DataFrame(data)

# Lista de valores para controle de frequência
valores_frequencia = ['P', 'A', 'J']


# Função para gerar controle de frequência aleatório
def gerar_frequencia_randomica(num_dias):
    return [random.choice(valores_frequencia) for _ in range(num_dias)]


# Função para criar ou acessar uma sheet de controle de frequência para um determinado mês
def criar_sheet_por_mes(sheet, nome_do_mes, funcionarios_df, num_dias):
    # Tenta abrir a aba do mês, se não existir, cria uma nova aba
    try:
        controle_frequencia_sheet = sheet.worksheet(nome_do_mes)
    except gspread.exceptions.WorksheetNotFound:
        controle_frequencia_sheet = sheet.add_worksheet(title=nome_do_mes, rows=len(funcionarios_df) + 1,
                                                        cols=num_dias + 2)

    # Criar um DataFrame para o controle de frequência
    frequencia_df = pd.DataFrame()
    frequencia_df['Nome do Funcionário'] = funcionarios_df['Nome do Funcionário']
    frequencia_df['Área'] = funcionarios_df['Área']

    # Preencher randomicamente a frequência para cada funcionário
    for dia in range(1, num_dias + 1):
        frequencia_df[f'Dia {dia:02d}'] = gerar_frequencia_randomica(len(funcionarios_df))

    # Atualizar a planilha de controle de frequência com os novos dados
    controle_frequencia_sheet.update([frequencia_df.columns.values.tolist()] + frequencia_df.values.tolist())


# Criar ou acessar múltiplas sheets de controle de frequência por mês
meses = {
    'Janeiro': 31,
    'Fevereiro': 28,
    'Março': 31,
    'Abril': 30,
    'Maio': 31,
    'Junho': 30,
    'Julho': 31,
    'Agosto': 31,
    'Setembro': 30,
    'Outubro': 31,
    'Novembro': 30,
    'Dezembro': 31
}

# Gerar controle de frequência para cada mês
for mes, dias in meses.items():
    criar_sheet_por_mes(sheet, mes, funcionarios_df, dias)

print("Controle de frequência para todos os meses criado com sucesso!")
