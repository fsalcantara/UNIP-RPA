 # Projeto de Automação de Controle de Frequência com Google Sheets e Gmail API
 
 Este projeto foi desenvolvido como parte do Trabalho de Conclusão de Curso (TCC) da **Universidade Paulista (UNIP)**, com foco na **Automação de Processos Robóticos (RPA)**. O objetivo principal é automatizar o controle de frequência de funcionários utilizando as APIs da Google para **Google Sheets** e **Gmail**, simplificando a coleta, análise e comunicação dos dados de presença.
 
 ## Objetivos do Projeto
 
 1. **Automatizar o controle de frequência**: Gerar, de forma automatizada, relatórios de presença e ausência de funcionários, preenchendo uma planilha no Google Sheets.
 2. **Analisar os dados de frequência**: Processar e analisar os dados de presença, gerando gráficos para auxiliar na visualização das informações.
 3. **Envio automático de relatórios**: Utilizar a API do Gmail para enviar os relatórios de frequência, com gráficos anexados, diretamente por e-mail.
 
 ## Requisitos do Sistema
 
 1. Credenciais de Serviço Google
 - Um arquivo de credenciais **JSON** para uma **Conta de Serviço** do Google Cloud é necessário para acessar as APIs do Google Sheets e Gmail.
 - A conta de serviço precisa ter permissões adequadas, como o acesso às planilhas do Google Sheets e a capacidade de enviar e-mails em nome de um usuário real, configurado no **Admin Console** do Google Workspace.
 
 ### 2. APIs Utilizadas
 - **Google Sheets API**: Utilizada para ler e atualizar os dados de frequência diretamente em planilhas hospedadas no Google Drive.
 - **Gmail API**: Utilizada para enviar e-mails automaticamente com os relatórios de frequência e gráficos.
 
 ## Instalação do Projeto

 1. Clone o repositório do projeto:
 
     ```hbash
     git clone https://github.com/seu-usuario/tcc-unip-rpa.git
     cd tcc-unip-rpa
     ```
 
 2. Instale as dependências do projeto:
 
     ```hbash
     pip install -r requirements.txt
     ```
 
 3. Coloque o arquivo `credenciais.json` no diretório principal do projeto. Esse arquivo será usado para autenticar o acesso às APIs da Google.
 
 ## Estrutura dos Scripts
 
 ### 1. `gerar_frequencia.py`
 
 Este script tem como função gerar automaticamente a frequência (presenças, ausências e justificativas) para os funcionários listados na planilha do Google Sheets.
 
 #### Como Executar:

 ```hbash
 python gerar_frequencia.py
 ```
 
 #### Funcionalidades:
 - Acessa uma planilha predefinida no Google Sheets e preenche, de forma automática, as presenças (`P`), ausências (`A`) e justificativas (`J`) para cada funcionário em um período de 30 dias.
 - Utiliza a API do Google Sheets para realizar leituras e escritas na planilha.
 
 ### 2. `analisar_frequencias.py`
 
 Este script analisa os dados de frequência gerados na planilha, cria gráficos para visualização de dados e envia os relatórios por e-mail para os responsáveis, utilizando a **Gmail API**.

 #### Como Executar:
 
 ```hbash
 python analisar_frequencias.py
 ```
 
 #### Funcionalidades:
 - **Análise de Dados**: Analisa os dados de presença e ausência e calcula as porcentagens de frequência.
 - **Geração de Gráficos**: Cria gráficos de barras para comparar a presença e a ausência de funcionários, além de gráficos de pizza para mostrar a distribuição total de presenças, ausências e justificativas.
 - **Envio Automático de Relatórios**: Envia um e-mail com o relatório completo e os gráficos anexados para o destinatário especificado, utilizando a API do Gmail.
 
 #### Relatórios e Gráficos:
 - Os gráficos gerados são armazenados na pasta `analise_grafica/` automaticamente após a execução do script.
 - Tipos de gráficos:
   1. **Gráfico de Barras**: Comparação entre presenças e ausências.
   2. **Gráfico de Pizza**: Distribuição total de presenças, ausências e justificativas.
 
 ## Configuração de Autenticação
 
 ### Google Sheets API
 - Certifique-se de que a **Google Sheets API** está ativada no Google Cloud Console e que o arquivo `credenciais.json` tem as permissões necessárias para acessar a planilha de controle de frequência.
 
 ### Gmail API
 - Certifique-se de que a **Gmail API** está ativada e que a conta de serviço configurada no arquivo `credenciais.json` tem a **delegação de autoridade** configurada para enviar e-mails em nome de um usuário real do Google Workspace.
 - A conta de serviço deve ter permissões adequadas no **Admin Console** do Google Workspace para o escopo `https://www.googleapis.com/auth/gmail.send`.
 
 ## Como Configurar a Conta de Serviço
 
 1. **Google Cloud Console**:
    - Crie uma **Conta de Serviço** e baixe o arquivo de credenciais em formato JSON.
    - Ative as APIs necessárias (Google Sheets API e Gmail API).
 
 2. **Google Workspace Admin Console**:
    - No **Admin Console**, vá até **Segurança** > **Configurações da API** e adicione o **ID da Conta de Serviço** para conceder permissões de delegação de domínio.
    - Autorize o escopo `https://www.googleapis.com/auth/gmail.send` para que a conta de serviço possa enviar e-mails em nome de um usuário.
 
 ## Estrutura do Projeto
 
 ```
 .
 ├── gerar_frequencia.py        # Script para gerar frequências automáticas
 ├── analisar_frequencias.py    # Script para analisar frequência e enviar relatórios
 ├── credenciais.json           # Credenciais da conta de serviço (não incluídas no repositório)
 ├── analise_grafica/           # Diretório onde os gráficos são salvos
 ├── README.md                  # Documentação do projeto
 ├── requirements.txt           # Dependências do projeto
 ```
 
 ## Requisitos de Software
 
 Listei as dependências necessárias no arquivo `requirements.txt`. Para instalar as dependências, use:
 
 ```hbash
 pip install -r requirements.txt
 ```
 
 Principais pacotes utilizados:
 - `gspread`: Acessa e atualiza dados nas planilhas do Google Sheets.
 - `google-auth`: Gerencia a autenticação via contas de serviço do Google.
 - `google-api-python-client`: Interage com as APIs da Google (Sheets e Gmail).
 - `matplotlib`: Gera gráficos para visualização dos dados.
 - `pandas`: Manipulação de dados tabulares.
 
 ## Contribuições
 
 Este projeto foi desenvolvido como parte de um Trabalho de Conclusão de Curso na **Universidade Paulista (UNIP)**, com foco em **RPA (Automação de Processos Robóticos)**. Contribuições e melhorias são bem-vindas. Se tiver sugestões ou identificar problemas, por favor, abra uma issue ou envie um pull request.
