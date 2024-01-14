import win32com.client as win32

def enviar_email_outlook(destinatario, assunto, corpo):
    # Criar a integração com o Outlook
    outlook = win32.Dispatch('outlook.application')

    # Criar um email
    email = outlook.CreateItem(0)

    # Configurar as informações do seu e-mail
    email.To = destinatario
    email.Subject = assunto
    email.HTMLBody = corpo

    # Enviar o e-mail
    email.Send()
    print("Email Enviado")

# Exemplo de uso da função
if __name__ == "__main__":
    destinatario = "espeditoa100@gmail.com"
    assunto = "Projeto Maromba - Resultados da Semana"
    corpo = f"""
    <p>Olá Janaina e Espedito!!</p>
    <p>Segue para o conhecimento de vocês os resultados de treino da semana.</p>

    <p>Abs,</p>
    <p>Código Python</p>
    """
    enviar_email_outlook(destinatario, assunto, corpo)
