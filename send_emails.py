import pandas as pd
import smtplib as sm
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from password import password

#Information
name_from = "Christian Carstens"
email_from = "ccarstens256@gmail.com"
file_name = "EmailList"
subject = "Test"


#Function that return an email object to send
def create_email(name_from, email_to):
    message = MIMEMultipart("alternative")
    message["Subject"] = "Test"
    message["From"] = f"{name_from} <{email_from}>"
    message["From"] = f"{name_to} <{email_to}>"
    html = f"""
    <html>
    <body>
        <p>Hola, {name_to}<br>
        Tus resultados son los siguientes:<br>
        </p>
        <h4>Nota Pregunta 1: 2<h4>
        <h4>Nota Pregunta 2: 7<h4>
    </body>
    </html>
    """
    part2 = MIMEText(html, "html")
    message.attach(part2)
    return message.as_string()

#Read excel
students_info = pd.read_excel(file_name + ".xlsx")
students_names = students_info['Nombre']
students_emails= students_info['Email']

#Server
server = sm.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(email_from, password)

#Create and send all emails
for name_to, email_to in zip(students_names, students_emails):
    try:
        server.sendmail(email_from, [email_to], create_email(name_to, email_to))
        print(f'Envío exitoso a {email_to}.')
    except Exception as error:
        print(f'Envío fallido a {email_to}.\n Se obtuvo el siguient error: {error}')

#Close the server
server.close()