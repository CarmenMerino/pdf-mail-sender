from PyPDF2 import PdfReader, PdfWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas, smtplib, ssl, numpy

FORM_SOURCE_PATH = "C:/Users/carme/OneDrive/Escritorio/PRONIED/prueba/constancia.pdf"
DATA_SOURCE_PATH = "C:/Users/carme/OneDrive/Escritorio/PRONIED/prueba/PRUEBA2.xlsx"

sender_email = input("Enter your email:")
password = input("Enter your password:")

reader = PdfReader(FORM_SOURCE_PATH)
writer = PdfWriter()

page = reader.pages[0]
fields = reader.get_fields()

writer.add_page(page)

excel_data = pandas.read_excel(DATA_SOURCE_PATH)
nombres = excel_data["nombre"]
codigos = excel_data["codigo"]
correos = excel_data["correo"]

for i in range(0, len(nombres)):
    nombre = nombres[i] if not nombres[i] is numpy.nan else '-'
    codigo = codigos[i] if not codigos[i] is numpy.nan else '-'
    receiver_email = correos[i]

    payload = {
        "nombre": nombre,
        "codigo": codigo
    }
    
    writer.update_page_form_field_values(writer.pages[0], payload)
    filename = f'constancia_{codigo}.pdf'
    # write "output" to PyPDF2-output.pdf
    with open(filename, "wb") as output_stream:
        writer.write(output_stream)
    
    if not receiver_email is numpy.nan:
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = "Constancia de participación taller ASITEC"
        message["Bcc"] = receiver_email  # Recommended for mass emails
        message["Cc"] = 'comunicacionesasitec@pronied.gob.pe'
        body = f'Buenas tardes {nombre}, gracias por haber participado en el taller de inducción de asistencia técnica del PRONIED. A continuación te adjuntamos tu certificado de participación. Te recomendamos usar el navegador Google Chrome para poder descargarlo.'
        # Add body to email
        message.attach(MIMEText(body, "plain"))

        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        # context = ssl.create_default_context()
        # with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        #     server.login(sender_email, password)
        #     server.sendmail(sender_email, receiver_email, text)


        # with smtplib.SMTP("smtp.office365.com", 587) as server:
        #     server.ehlo()
        #     server.starttls()
        #     server.login(sender_email, password)
        #     text = message.as_string()
        #     server.sendmail(sender_email, receiver_email, text)
        #     print('email sent')
        #     server.quit()


        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, password)
            text = message.as_string()
            server.sendmail(sender_email, receiver_email, text)
            print('email sent')
            server.quit()