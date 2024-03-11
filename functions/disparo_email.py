import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def enviar_email(file_path, file_name):
    # Configurações do servidor SMTP
    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    smtp_username = 'pedrolopesv@outlook.com'
    smtp_password = ''

    # Configurações do e-mail
    sender_email = 'pedrolopesv@outlook.com'
    receiver_email = 'pedrolopesv@outlook.com'
    subject = 'TESTE ENVIO DO EXCEL'
    body = 'TESTE ENVIO DO EXCEL'

 # Criando a mensagem multipart
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Adicionando o corpo do e-mail
    message.attach(MIMEText(body, 'plain'))

  # Anexando o arquivo Excel
    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {file_name}')
        message.attach(part)

    # Iniciando a conexão SMTP e enviando o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    print("E-mail enviado com sucesso!")
