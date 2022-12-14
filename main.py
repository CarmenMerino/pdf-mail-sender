from PyPDF2 import PdfReader, PdfWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas, smtplib, ssl, numpy

FORM_SOURCE_PATH = "D:/test_form_2.pdf"
DATA_SOURCE_PATH = "D:/data_test.xlsx"

sender_email = input("Enter your email:")
password = input("Enter your password:")

reader = PdfReader(FORM_SOURCE_PATH)
writer = PdfWriter()

page = reader.pages[0]
fields = reader.get_fields()

writer.add_page(page)

excel_data = pandas.read_excel(DATA_SOURCE_PATH)
companies = excel_data["txt_company"]
codes = excel_data["txt_code"]
emails = excel_data["txt_mail"]

for i in range(0, len(companies)):
    company = companies[i] if not companies[i] is numpy.nan else '-'
    code = codes[i] if not codes[i] is numpy.nan else '-'
    receiver_email = emails[i]

    payload = {
        "txt_company": company,
        "txt_code": code
    }
    
    writer.update_page_form_field_values(writer.pages[0], payload)
    filename = f'D:/filled-{company}.pdf'
    # write "output" to PyPDF2-output.pdf
    with open(filename, "wb") as output_stream:
        writer.write(output_stream)
    
    if not receiver_email is numpy.nan:
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = "Test invoice email"
        message["Bcc"] = receiver_email  # Recommended for mass emails
        body = "This is an email with attachment sent from Python"
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