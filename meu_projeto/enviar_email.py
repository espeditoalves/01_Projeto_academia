import win32com.client as win32
from scripts.functions import *


def enviar_email_outlook(
    destinatarios, assunto, dataframe_anexo, imagem1, imagem2
):
    # Criar a integração com o Outlook
    outlook = win32.Dispatch('outlook.application')

    # Criar um email
    email = outlook.CreateItem(0)

    # Configurar as informações do seu e-mail
    # Junta os destinatários com ponto e vírgula
    email.To = ';'.join(destinatarios)
    email.Subject = assunto

    # Começo do e-mail
    corpo_email = f"""
    <p>Olá Janaina e Espedito!!</p>
    <p>Segue para o conhecimento de vocês os resultados de treino da semana.</p>
    """
    # Adicionar o DataFrame ao corpo do e-mail
    corpo_email += dataframe_anexo.to_html(index=False)
    # Adiciona o Desfecho
    corpo_email += f"""
    <p>Abs,</p>
    <p>Desenvolvedor: Espedito Ferreira Alves</p>
    """
    email.HTMLBody = corpo_email

    # Adicionar imagens como anexos
    email.Attachments.Add(imagem1)
    email.Attachments.Add(imagem2)
    # Enviar o e-mail
    email.Send()
    print('Email Enviado')


# A estrutura __main__ é geralmente usada para determinar se o script está sendo executado como um programa independente
# ou se está sendo importado como um módulo em outro script.

# Dessa forma, quando você executa o script, o bloco if __name__ == "__main__": será verdadeiro,
# e a função enviar_email_outlook será chamada.
# Se você importar este script como um módulo em outro lugar, o bloco if __name__ == "__main__": será falso,
# e a função não será executada automaticamente.
if __name__ == '__main__':
    # Se o script estiver sendo executado como um programa independente

    # Configuração dos destinatários, assunto, etc.
    destinatarios = ['espeditoa100@gmail.com']  # 'janainavdm@gmail.com']
    assunto = 'Projeto Treino Firme - Resultados da Semana'
    dataframe_anexo, data_formatada = analisa_excel()
    imagem1 = f'C:/Users/esped/OneDrive/Documentos/_repositorios_/Projetos/01_Projeto_academia/meu_projeto/graficos/{data_formatada}_grafico_seaborn_Espedito.png'
    imagem2 = f'C:/Users/esped/OneDrive/Documentos/_repositorios_/Projetos/01_Projeto_academia/meu_projeto/graficos/{data_formatada}_grafico_seaborn_Janaina.png'

    # Chama a função principal
    enviar_email_outlook(
        destinatarios, assunto, dataframe_anexo, imagem1, imagem2
    )
