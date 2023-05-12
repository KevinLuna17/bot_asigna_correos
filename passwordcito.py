import poplib
import email
import smtplib

# Configurar conexión POP3
pop_conn = poplib.POP3_SSL('outlook.office365.com')
pop_conn.user('correo@hotmail.com')
pop_conn.pass_('pass')

# Obtener lista de destinatarios
destinatarios = ['destinatario1@example.com', 'destinatario2@example.com', 'destinatario3@example.com']

# Obtener número de correos electrónicos en la bandeja de entrada
num_mensajes = len(pop_conn.list()[1])

# Asignar cada correo electrónico al siguiente destinatario en la lista
for i in range(num_mensajes):
    # Obtener correo electrónico y analizarlo
    _, mensaje_raw = pop_conn.retr(i+1)
    mensaje = email.message_from_bytes(b'\r\n'.join(mensaje_raw))
    
    # Extraer contenido y destinatarios del correo electrónico
    contenido = mensaje.get_payload()
    destinatario = destinatarios[i % len(destinatarios)]
    
    # Configurar conexión SMTP y enviar correo electrónico al destinatario asignado
    smtp_conn = smtplib.SMTP('smtp.office365.com', 587')
    smtp_conn.login('correo@hotmail.com', 'pass')
    smtp_conn.sendmail('correo@hotmail.com', destinatario, contenido)
    smtp_conn.quit()
    
    # Marcar correo electrónico como leído
    pop_conn.dele(i+1)

# Cerrar conexión POP3
pop_conn.quit()
