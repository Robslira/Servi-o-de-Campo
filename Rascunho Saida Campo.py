import pandas as pd
import pywhatkit as kit
import pyautogui
import time
from datetime import datetime

# Exemplo de horário
horario = "08:30:00"

# Converter o horário para o formato desejado
horario_formatado = datetime.strptime(horario, "%H:%M:%S").strftime("%H:%M")

print(horario_formatado)  # Saída: 08:45

# Caminho do arquivo Excel
caminho_arquivo = "Saida_Campo.xlsx"

# Ler o arquivo Excel e remover espaços extras dos nomes das colunas
df = pd.read_excel(caminho_arquivo, sheet_name="Planilha1")
df.columns = df.columns.str.strip()


# Lista de destinatários que não devem receber lembretes
ignorar_destinatarios = [ "Maria Alcantara", "Anizio Arruda", "America Ferreira", "Rosalina Jesus",
                          "Sinomar Marcelino", "Genoveva Miranda", "Luzia Oliveira", "Maria Oliveira",                         
                          "Deyvison Silva", "Kalleb Soares", "Edna Silva","Diego Souza", "Francinete Moura",
                          "Gabrielly Souza", "Rubia Santos", "Thais Soares", "Fernanda Lima", "Rian Moraes",
                          "Sabrina Magalhaes", "Divina Damiao","Elaine Martins","Emily Silva","Lucilene Rosa",                          "Samyrah Oliveira","Sthefany Martins","Dayana Alves","Rejane Arantes","Rennan Monteiro",
                          "Rosa Marcelino", "Thais Souza", "Leni Miranda","Luciana Rodrigues","Marco Aurelio",
                          "Sarah Oliveira","Adriele Barreto","Alzerino Santos","Antonio Rosa","Dennis Damiao",
                          "Ivo Soares","Neuza Bastos","Rejane Arantes","Rennan Monteiro"]

                          
# Exibir as primeiras linhas do DataFrame para verificar os dados
# Listar as colunas disponíveis para verificar a estrutura do DataFrame
print("Colunas do DataFrame após ajustes:", df.columns.tolist())

# Converter as colunas de data para o formato datetime e tratar erros
df['data1'] = pd.to_datetime(df['data1'], format="%d/%m/%Y", errors='coerce')
df['data2'] = pd.to_datetime(df['data2'], format="%d/%m/%Y", errors='coerce')

# Dicionário para traduzir os dias da semana
dias_semana = { 
                    "Monday": "segunda-feira",
                    "Tuesday": "terça-feira",
                    "Wednesday": "quarta-feira",
                    "Thursday": "quinta-feira",
                    "Friday": "sexta-feira",
                    "Saturday": "sábado",
                    "Sunday": "domingo"}

# Função para ajustar o formato do número de telefone
def ajustar_telefone(numero):
    if pd.isna(numero):  # Verifica se o valor é nulo ou NaN
        return None  # Retorna None para valores ausentes
    numero = str(numero).replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    if not numero.startswith("+55"):  # Adicionar o código do Brasil se não estiver presente
        return f"+55{numero}"
    return numero

# Aplicar o ajuste de telefone na coluna 'celular'
df['celular'] = df['celular'].apply(ajustar_telefone)

# Iterar pelas linhas do DataFrame e processar os dados
for _, row in df.iterrows():
    try:
        # Ler o perfil e garantir valor padrão
        perfil = row.get("Perfil", None)
        if pd.isna(perfil) or perfil not in ["Irmão", "Irmã", "Estudante"]:
            perfil = "Perfil não especificado"

        # Ler o nome e número de celular
        nome = row.get("nome", "Participante")
        celular = row.get("celular", None)

        # Pular caso o nome esteja na lista de destinatários a serem ignorados
        if nome in ignorar_destinatarios:
            print(f"Lembrete não será enviado para {nome}. Pulando.")
            continue

        # Pular caso o número de celular esteja ausente
        if not celular:
            print(f"Número de celular ausente para {perfil} {nome}. Pulando.")
            continue

        # Obter informações do primeiro local
        if not pd.isna(row['data1']):
            dia_semana1 = row['data1'].strftime("%A")
            dia_semana_pt1 = dias_semana.get(dia_semana1, "Dia não especificado")
            data1 = row['data1'].strftime(f"%d/%m/%Y ({dia_semana_pt1})")
            local1 = row.get('local1', 'Local não especificado')
            endereco1 = row.get('endereço1', 'Endereço não informado')
            bairro1 = row.get('bairro1', 'Bairro não especificado')
            horario1 = row.get('horario1', 'Horário não especificado')
            campanha1 = row.get('campanha1', 'Campanha não especificada')
        else:
            data1, local1, endereco1, bairro1, horario1, campanha1 = [None] * 6

        # Obter informações do segundo local
        if not pd.isna(row['data2']):
            dia_semana2 = row['data2'].strftime("%A")
            dia_semana_pt2 = dias_semana.get(dia_semana2, "Dia não especificado")
            data2 = row['data2'].strftime(f"%d/%m/%Y ({dia_semana_pt2})")
            local2 = row.get('local2', 'Local não especificado')
            endereco2 = row.get('endereço2', 'Endereço não informado')
            bairro2 = row.get('bairro2', 'Bairro não especificado')
            horario2 = row.get('horario2', 'Horário não especificado')
            campanha2 = row.get('campanha2', 'Campanha não especificada')
        else:
            data2, local2, endereco2, bairro2, horario2, campanha2 = [None] * 6

        # Criar a mensagem consolidada com 'Perfil' antes do nome
        mensagem = f"Olá {perfil},\n {nome},\ntudo bem? 😊\n\nEste é um lembrete automático para informar\nos detalhes da sua saída de campo:\n\n"
        
        # Adicionar informações do primeiro local
        if data1:
            mensagem += f"📍 Local 1: {local1}\n🏘️ Endereço: {endereco1}\n📍 Bairro: {bairro1}\n\n📅 Data: {data1}\n⏰ Horário: {horario1}\n📢 Campanha: {campanha1} "

        # Adicionar informações do segundo local
        if data2:
            mensagem += f"📍 Local 2: {local2}\n🏘️ Endereço: {endereco2}\n📍 Bairro: {bairro2}\n\n📅 Data: {data2}\n⏰ Horário: {horario2}\n📢 Campanha: {campanha2} "

        # Criar fechamento da mensagem
        mensagem += "😊"

        # Enviar mensagem via WhatsApp
        print(f"Enviando mensagem para {perfil} {nome} ({celular})...")
        kit.sendwhatmsg_instantly(celular, mensagem)
        print(f"Mensagem enviada para {perfil} {nome} ({celular}).")

        # Pausar para evitar erros e fechar a aba do WhatsApp
        time.sleep(5)  # Garantir tempo suficiente para envio
        pyautogui.hotkey('ctrl', 'w')  # Fecha a aba aberta
        print(f"Aba do WhatsApp fechada para {perfil} {nome}.")
        
    except Exception as e:
        print(f"Erro ao processar ou enviar mensagem para {perfil} {nome} ({celular}): {e}")