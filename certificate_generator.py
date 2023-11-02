from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import yagmail
from config import pass_gmail ############################# ---- CONFIGURE SUA SENHA ------------

class EditCertificate:
    '''classe que cria certificado dinamicamente e envia por email'''
    def __init__(self):
        self.read_data()
        self.send_email_with_certificate()

    def read_data(self):
        self.df = pd.read_excel('lista_alunos.xlsx')

    def send_email_with_certificate(self):
        for _, row in self.df.iterrows():
            name = row['Nome']  
            email = row['Email']  
            
            # Gerar um certificado com o nome
            certificate_image = self.create_certificate(name)

            # Enviar email personalizado
            self.send_email_generic(name, email, certificate_image)

    def create_certificate(self, name):
        # Abrir a imagem de modelo
        imagem = Image.open('template.png')
        
        draw = ImageDraw.Draw(imagem)

        font = ImageFont.truetype('calibrib.ttf', 150)

        # Definindo a cor da fonte como preto (0, 0, 0)
        draw.text((625, 950), name, font=font, fill=(0, 0, 0))

        certificate_image = f'{name}_Certificate_Evangelizando_SQL.png'
        imagem.save(certificate_image)
        return certificate_image

    def send_email_generic(self, name, email, certificate_image):
        # Inicializar o Yagmail SMTP
        usuario = yagmail.SMTP(user='seu_email@gmail.com', password=pass_gmail) ############################# ---- ALTERE POR SEU EMAIL ------------
        
        assunto = 'Certificado de Participação - Evangelizando o SQL'
        conteudo = f'Olá {name},\n\nAqui está o seu segundo certificado, pois verifiquei que o primeiro não ficou bom. \n\nAtenciosamente,\nAdministração Evangelizando SQL'

        # Anexar a imagem do certificado
        usuario.send(
            to=email,
            subject=assunto,
            contents=conteudo,
            attachments=certificate_image
        )

        print(f'Email enviado para {email} com sucesso!')

start = EditCertificate()
