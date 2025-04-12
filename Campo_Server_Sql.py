import pywhatkit as kit
import pyodbc
import time

# Configuração de conexão ao SQL Server
conexao = pyodbc.connect(
    'DRIVER={SQL Server};SERVER=SEU_SERVIDOR;DATABASE=CONGREGACAO_GUARANI_88062;UID=USUARIO;PWD=SENHA'
)
cursor = conexao.cursor()

# Consulta para extrair informações do destinatário específico
query = """
SELECT NOME, ENDERECO, GRUPO
FROM FOR_CAD_PUBL
WHERE DATA = '2025-04-13'
  AND HORARIO = '08:45'
  AND GRUPO <> '0'
  AND CELULAR = '55 34 996853975';
"""
cursor.execute(query)
resultados = cursor.fetchall()

# Iterar sobre os resultados e enviar a mensagem
for row in resultados:
    try:
        nome = row[0]
        endereco = row[1]
        grupo = row[2]
        celular = '34 996853975'  # Garantir que apenas esse número receba o lembrete

        mensagem = f"""Olá {nome},
Tudo bem? 😊

Este é um lembrete automático para informar:
📅 Data: 13/04/2025
⏰ Horário: 08:45
🏠 Endereço: {endereco}
🏘️ Local: Residência do irmão Roberto Lira

Grupo: {grupo}
😊"""

        # Enviar mensagem pelo WhatsApp
        print(f"Enviando mensagem para {nome} ({celular})...")
        kit.sendwhatmsg_instantly(celular, mensagem)

        # Pausar para evitar problemas
        time.sleep(5)
        print(f"Mensagem enviada para {nome} ({celular}).")

    except Exception as e:
        print(f"Erro ao enviar mensagem para {nome}: {e}")

# Fechar conexão
cursor.close()
conexao.close()

