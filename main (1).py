import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import smtplib
import time

contatos = pd.read_excel("Lista que tem email.xlsx")

# Inicializa a lista de e-mails enviados
emails_enviados = []
emails_rejeitados = []

# Configuração do servidor SMTP
smtp_server = 'smtp.terra.com.br'
smtp_port = 587
smtp_username = 'exemplo@ex.com.br'
smtp_password = '***********'

# Tamanho do lote
lote_size = 100

# Número total de lotes
total_lotes = len(contatos) // lote_size + (len(contatos) % lote_size > 0)

# Loop para enviar os e-mails em lotes
for i in range(0, len(contatos), lote_size):
    print(f"Enviando lote {i // lote_size + 1} de {total_lotes}")

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)

        for index, cliente in contatos.iloc[i:i+lote_size].iterrows():
            try:
                msg = MIMEMultipart()
                msg['Subject'] = 'Email colégio Vaz'
                msg['From'] = 'exemplo@ex.com.br'
                msg['To'] = cliente['Email']

                message = f" convida todos para o Festival da Freguesia do Ó"

                # Adicione a imagem no corpo do email usando a tag HTML
                with open('Image.jpeg', 'rb') as file:
                    image = MIMEImage(file.read())
                    image.add_header('Content-ID', '<image1>')
                    msg.attach(image)
                message += '<br><img src="cid:image1" width="400">'

                msg.attach(MIMEText(message, 'html'))

                # Adicione o arquivo PDF como anexo
                with open('exemplopdf', 'rb') as file:
                    pdf = MIMEApplication(file.read(), _subtype='pdf')
                    pdf.add_header('Content-Disposition', 'attachment', filename='exemplo.pdf')
                    msg.attach(pdf)

                server.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

                # Adiciona o e-mail à lista de e-mails enviados
                emails_enviados.append(cliente['Email'])
                print(f'Enviada com sucesso para {cliente["Email"]}')
            except smtplib.SMTPRecipientsRefused as e:
                # Caso ocorra um erro de rejeição do destinatário, pula para o próximo
                print(f'Erro no envio para {cliente["Email"]}: {e}')
                emails_rejeitados.append(cliente['Email'])
            except smtplib.SMTPServerDisconnected as e:
                # Caso o servidor SMTP tenha desconectado, reconecte e continue o envio
                print(f'Erro de desconexão do servidor SMTP: {e}')
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(smtp_username, smtp_password)
                continue


        # Intervalo de 10 segundos entre os lotes
        time.sleep(10)
    except smtplib.SMTPServerDisconnected as e:
        # Caso o servidor SMTP tenha desconectado durante o lote, reconecte e continue o envio
        print(f'Erro de desconexão do servidor SMTP durante o lote: {e}')
        continue
    except smtplib.SMTPDataError as e:
        print("Sua cota estourou Tente daqui 24h")
        
 # Salva a lista de e-mails enviados em um arquivo de texto
    with open('emails_enviados.txt', 'a') as file:
        for email in emails_enviados:
            file.write(email + '\n')

        # Salva a lista de e-mails rejeitados em um arquivo de texto
    with open('emails_rejeitados.txt', 'a') as file:
        for email in emails_rejeitados:
            file.write(email + '\n')


print("Operação finalizada")