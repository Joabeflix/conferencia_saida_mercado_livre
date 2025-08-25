import os
import smtplib
from email.mime.text import MIMEText
from utils.utils import texto_no_console
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class Email:
    def __init__(self, remetente: str) -> None:
        self.remetente = remetente
        self.senha: str | None = None

    def definir_senha(self, senha: str):
        self.senha = senha

    def enviar(self, assunto: str, mensagem: str, destinatarios: str, anexos: list | None=None):

        msg = MIMEMultipart()
        msg['Subject'] = assunto
        msg['From'] = self.remetente
        msg['To'] = ', '.join(destinatarios)

        corpo = MIMEText(mensagem, 'plain')
        msg.attach(corpo)

        if anexos:
            for caminho_arquivo in anexos:
                if os.path.exists(caminho_arquivo):
                    with open(caminho_arquivo, 'rb') as f:
                        parte = MIMEApplication(f.read(), Name=os.path.basename(caminho_arquivo))
                        parte['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_arquivo)}"'
                        msg.attach(parte)
                else:
                    texto_no_console(f"Atenção: Arquivo '{caminho_arquivo}' não encontrado e não foi anexado.")

        # Envia o e-mail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
            if self.senha:
                smtp_server.login(self.remetente, self.senha)
                smtp_server.sendmail(self.remetente, destinatarios, msg.as_string())
        texto_no_console("Email enviado com sucesso.")


# def fazer_envio_email():
#     email = Email(
#         remetente='joabealvesyt@gmail.com',
#     )
#     email.definir_senha('jxkf ygym pedj dxso')

#     hostname = socket.gethostname()
#     ip = socket.getaddrinfo(hostname, None, family=socket.AF_INET)[0][4][0]


#     current_datetime = datetime.now()

#     formatted_datetime = current_datetime.strftime("%d-%m-%Y %H:%M:%S")

#     _body = f"""
#     Atualização da planilha do GoogleSheets enviada.\n
#     EMPRESA: {'teste'}
#     USUÁRIO: {os.environ['USERNAME']}
#     NOME DO COMPUTADOR: {os.environ['COMPUTERNAME']}
#     IP DO COMPUTADOR: {ip}
#     NOME DO DOMÍNIO: {os.environ['USERDOMAIN']}
#     DNS: {os.environ['USERDNSDOMAIN']}
#     DATA E HORA DO ENVIO: {formatted_datetime}
#     """

#     email.enviar(assunto=f'PLANILHA ENVIADA NA AUTOMAÇÃO [{'teste'}]', mensagem='Olá', destinatarios=['joabe@samarc.com.br'], anexos=[r'temp\teste.txt'])
