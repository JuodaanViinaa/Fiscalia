import datetime
import smtplib
# MIMEMultipart send emails with both text content and attachments.
from email.mime.multipart import MIMEMultipart
# MIMEText for creating body of the email message.
from email.mime.text import MIMEText
# MIMEApplication attaching application-specific data (like CSV files) to email messages.
from email.mime.application import MIMEApplication

def enviar_mensaje(destinatario = "maldonadodaniel96@outlook.com"):
    """
    Funcion para enviar por correo electronico los dos archivos de resumen: Informe ejecutivo y Audiencias
    :param destinatario: La direccion de correo hacia la que se enviaran los archivos. Por defecto se envian a
    maldonadodaniel96@outlook.com
    :return:
    """
    subject = "Reporte Ejecutivo CGIDGAV"
    body = ""
    sender_email = "maldonadopythontest@gmail.com"
    recipient_email = destinatario
    sender_password = "axlz ljtn jzfr uclb "
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    path_to_file1 = (f'/home/daniel/PycharmProjects/Fiscalia/informeEjecutivo/{datetime.date.today().strftime("%Y%m%d")}/'
                     f'Informe diario/{datetime.datetime.now().strftime("%Y%m%d")} INFORME EJECUTIVO DIARIO CGIDGAV.xlsx')
    path_to_file2 = (f'/home/daniel/PycharmProjects/Fiscalia/informeEjecutivo/{datetime.date.today().strftime("%Y%m%d")}/'
                     f'Audiencias/{datetime.datetime.now().strftime("%Y%m%d")} Audiencias.xlsx')

    # MIMEMultipart() creates a container for an email message that can hold
    # different parts, like text and attachments and in next line we are
    # attaching different parts to email container like subject and others.
    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = recipient_email
    body_part = MIMEText(body)
    message.attach(body_part)

    # section 1 to attach file
    with open(path_to_file1,'rb') as file:
        # Attach the file with filename to the email
        message.attach(MIMEApplication(file.read(), Name=f'{datetime.date.today().strftime("%Y%m%d")} INFORMEEJECUTIVO DIARIO CGDIDGAV.xlsx'))
    with open(path_to_file2,'rb') as file:
        # Attach the file with filename to the email
        message.attach(MIMEApplication(file.read(), Name=f'{datetime.date.today().strftime("%Y%m%d")} Audiencias.xlsx'))

    # secction 2 for sending email
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
       server.login(sender_email, sender_password)
       server.sendmail(sender_email, recipient_email, message.as_string())