import win32com.client as win32
from functions import *

def enviar_email_outlook(destinatarios, assunto, dataframe_anexo, imagem1, imagem2):
    # Criar a integração com o Outlook
    outlook = win32.Dispatch('outlook.application')

    # Criar um email
    email = outlook.CreateItem(0)

    # Configurar as informações do seu e-mail
    email.To = ";".join(destinatarios)  # Junta os destinatários com ponto e vírgula
    email.Subject = assunto
    
    # Começo do e-mail
    corpo_email = f"""
    <p>Olá Janaina e Espedito!!</p>
    <p>Segue para o conhecimento de vocês os resultados de treino da semana.</p>
    """
    # Adicionar o DataFrame ao corpo do e-mail
    corpo_email += dataframe_anexo.to_html(index=False)
    # Adiciona o Desfecho
    corpo_email +=f"""
    <p>Abs,</p>
    <p>Desenvolvedor: Espedito Ferreira Alves</p>
    """
    email.HTMLBody = corpo_email

    # Adicionar imagens como anexos
    email.Attachments.Add(imagem1)
    email.Attachments.Add(imagem2)
    # Enviar o e-mail
    email.Send()
    print("Email Enviado")

destinatarios = ["espeditoa100@gmail.com", "janainavdm@gmail.com"]
assunto = "Projeto Treino Firme - Resultados da Semana"
dataframe_anexo, data_formatada = analisa_excel()
imagem1 = f'C:/Users/esped/OneDrive/Documentos/_repositorios_/Projetos/01_Projeto_academia/graficos/{data_formatada}_grafico_seaborn_Espedito.png'
imagem2 = f'C:/Users/esped/OneDrive/Documentos/_repositorios_/Projetos/01_Projeto_academia/graficos/{data_formatada}_grafico_seaborn_Janaina.png'

enviar_email_outlook(destinatarios, assunto, dataframe_anexo, imagem1, imagem2)