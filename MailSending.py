import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


def send_mail(send_from, send_to,send_cc,send_bcc, subject, message, files=[],
              server="localhost", port=587, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from

    if len(send_to)==1:
        msg['To'] = send_to
    else:
        msg['To'] =", ".join(send_to)

    if len(send_cc)==1:
        msg['Cc'] = send_cc
    else:
        msg['Cc']=", ".join(send_cc)

    if len(send_cc)==1:
        msg['Bcc']=send_bcc
    else:
        msg['Bcc']=", ".join(send_bcc)

    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

    print("Mail sent successfully")



send_mail(send_from="nominations@statmark.de",
          send_to=["lukas.dicke@web.de","lukas.dicke@statkraft.de"],
          send_cc= [],
          send_bcc= [],
          subject="TestPythonTestEmailFromStatmarkAccount",
          message="Automated email from Python script",
          files=[],
          server= "mail.123domain.eu",
          port=587,
          username="nominations@statmark.de",
          password="fk390krvadf",
          use_tls=True)

