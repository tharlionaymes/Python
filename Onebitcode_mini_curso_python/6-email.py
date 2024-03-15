import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# 1 - Dados do E-mail
password = open("senha","r").read()
from_email = "tharlionaymes@gmail.com"
to_email = "tharlionaymes@gmail.com"
subject = "Automação Planilha"
body = """
Teste de Email feito pelo Python
"""

# 2 - Montando a estrutura do E-mail
message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

message.set_content(body)
safe = ssl.create_default_context() #sem isso o gmail não envia

# 3 - Adicionar Anexo
anexo = "test.xlsx"
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].split("/")
with open(anexo,"rb") as a:
    message.add_attachment(
        a.read(),
        maintype=mime_type,
        subtype = mime_subtype,
        filename = anexo
    )

# 4 - Envio do E-mail
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as smtp:
    smtp.login(from_email, password)
    smtp.sendmail(
        from_email,
        to_email,
        message.as_string()
    )