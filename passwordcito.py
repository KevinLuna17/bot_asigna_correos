import poplib
import email
from email.header import decode_header
import smtplib

# Configurar conexión POP3
pop_conn = poplib.POP3_SSL('outlook.office365.com')
pop_conn.user('buzon@hotmail.com')
pop_conn.pass_('password')

# Obtener lista de destinatarios
destinatarios = ['destinatario1@hotmail.com', 'destinatario2@hotmail.com', 'destinatario3@outlook.es']
print(f'Destinatarios: {destinatarios}')

# Obtener información sobre los correos electrónicos en la cuenta
num_messages = len(pop_conn.list()[1])
print(f'Hay {num_messages} correos electrónicos en la cuenta')

# Recorrer los mensajes de la bandeja de entrada
for i in range(num_messages):
    # Obtener el número de mensaje
    message_num = i + 1
    destinatario = destinatarios[i % len(destinatarios)]
    # Recuperar el mensaje de correo electrónico
    _, message_lines, _ = pop_conn.retr(message_num)
    
    # Convertir las líneas del mensaje en una cadena
    message_content = b'\r\n'.join(message_lines)
    
    # Parsear el mensaje de correo electrónico
    parsed_message = email.message_from_bytes(message_content)
    
    # Obtener los encabezados del mensaje
    subject = parsed_message['Subject']
    from_address = parsed_message['From']
    to_address = parsed_message['To']
    
    # Decodificar los encabezados si es necesario
    decoded_subject = decode_header(subject)[0][0]
    decoded_from_address = decode_header(from_address)[0][0]
    decoded_to_address = decode_header(to_address)[0][0]
    
    # Imprimir los detalles del mensaje
    print(f"Mensaje {message_num}")
    print(f"De: {decoded_from_address}")
    print(f"Para: {decoded_to_address}")
    print(f"Asunto: {decoded_subject}")
    print("Cuerpo del mensaje:")
    
    # Obtener el cuerpo del mensaje
    body = ""
    if parsed_message.is_multipart():
        for part in parsed_message.walk():
            content_type = part.get_content_type()
            if content_type == 'text/plain':
                body = part.get_payload(decode=True).decode(errors='ignore')
                break
    else:
        body = parsed_message.get_payload(decode=True).decode(errors='ignore')
    
    print(body)
    print("-------------------")
    
    # Configuración del servidor SMTP de Microsoft
    smtp_server = 'smtp.office365.com'
    smtp_port = 587

    # Información de la cuenta de correo electrónico
    username = 'buzon@hotmail.com'
    password = 'password'

    # Crear conexión segura con el servidor SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()

    # Iniciar sesión en la cuenta de correo electrónico
    server.login(username, password)

    # Enviar correo electrónico
    from_address = 'buzon@hotmail.com'
    #to_address = decoded_to_address
    subject = subject
    cuerpo = body

    message = f"From: {from_address}\nTo: {destinatario}\nSubject: {subject}\n\n{cuerpo}"

    server.sendmail(from_address, destinatario, message)

    # Cerrar la conexión con el servidor SMTP
    server.quit()

# Cerrar la conexión al servidor POP3 de Outlook
pop_conn.quit()
