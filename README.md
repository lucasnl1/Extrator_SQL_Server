# Projeto criado para relatórios mensais na empresa em que estou
#Automação de Relatórios e Envio de Email

Este script Python realiza uma série de tarefas automatizadas que incluem:
1. Conectar-se a um banco de dados SQL Server e executar uma consulta SQL para recuperar dados.
2. Salvar os dados recuperados em um arquivo Excel formatado.
3. Enviar o arquivo Excel por email usando um servidor SMTP.

## Requisitos

Para executar este script, você precisará das seguintes bibliotecas Python:
- `pyodbc`: Para conectar-se ao banco de dados SQL Server.
- `pandas`: Para manipulação e salvamento de dados.
- `openpyxl`: Para manipulação do arquivo Excel.
- `smtplib`, `email`: Para envio de email com anexo.

Você pode instalar as bibliotecas necessárias utilizando o `pip`:
```sh
pip install pyodbc pandas openpyxl
