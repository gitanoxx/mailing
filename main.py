# Importamos las librerías necesarias
import smtplib
import pandas as pd

# ==============================
# CONFIGURACIÓN DE LA CUENTA
# ==============================

# Nombre que aparecerá como remitente
name_account = "Nombre de tu servicio"

# Correo del remitente (reemplazar por variable de entorno en producción)
email_account = "tu_email@ejemplo.com"

# Contraseña o App Password (NUNCA subir credenciales reales a repositorios)
password_account = "TU_PASSWORD_AQUI"

# ==============================
# CONEXIÓN AL SERVIDOR SMTP
# ==============================

# Configuración para Gmail (puede cambiar según proveedor)
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)

# Identificación con el servidor
server.ehlo()

# Inicio de sesión
server.login(email_account, password_account)

# ==============================
# LECTURA DEL ARCHIVO EXCEL
# ==============================

# Se lee el archivo Excel que contiene:
# Name | Email | Subject | Message
email_df = pd.read_excel("Data/Emails.xlsx")

# Se extraen las columnas necesarias
all_names = email_df['Name']
all_emails = email_df['Email']
all_subjects = email_df['Subject']
all_messages = email_df['Message']

# ==============================
# ENVÍO DE CORREOS
# ==============================

# Se recorre cada fila del archivo
for i in range(len(email_df)):

    # Datos del destinatario
    name = all_names[i]
    email = all_emails[i]

    # Personalización del asunto
    subject = all_subjects[i] + ', ' + name + '!'

    # Personalización del mensaje
    message = (
        'Hola, ' + name + '!\n\n' +
        all_messages[i] + '\n\n' +
        'Saludos,\n' +
        name_account
    )

    # Construcción del correo en formato texto plano
    sent_email = (
        "From: {0} <{1}>\n"
        "To: {2} <{3}>\n"
        "Subject: {4}\n\n"
        "{5}"
    ).format(name_account, email_account, name, email, subject, message)

    # Intento de envío del correo
    try:
        server.sendmail(email_account, [email], sent_email)
        print(f"Correo enviado a {email}")
    except Exception as e:
        print(f"No se pudo enviar el correo a {email}. Error: {e}")

# ==============================
# CIERRE DEL SERVIDOR
# ==============================

server.close()