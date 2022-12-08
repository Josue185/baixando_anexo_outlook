#imports das bibliotecas usadas

import imaplib
import email
import os
import base64

#Conectando ao Servidor do outlook
objCon = imaplib.IMAP4_SSL("outlook.office365.com")

#passando o login e senha para entrar no e-mail
login = "SeuEmail"
senha = "suaSenha"

objCon.login(login,senha)

#loopando a caixa de entrada
print(objCon.list())
objCon.select(mailbox='inbox',readonly=True)
respostas,idDosEmails = objCon.search(None,'All')

print(idDosEmails)

#Loopando cada id de email na caixa de entrada
for num in idDosEmails[0].split():
    #decodificando
    resutado,dados = objCon.fetch(num,'(RFC822)')
    texto_do_email = dados[0][1]
    texto_do_email = texto_do_email.decode('utf-8')# usando o decodificador para o texto do email
    texto_do_email = email.message_from_string(texto_do_email)# usando o formato string
    
    #loopando as partes dos emails
    for part in texto_do_email.walk():
        # Caso tenha anexo, pegar o nome.
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        
        #pegamos o anexo
        fileName = part.get_filename()
        #criamos um arquivo com o mesmo nome do arquivo
        arquivo = open(fileName, 'wb')
        #Criamos o binario do arquivo
        arquivo.write(part.get_payload(decode=True))
        arquivo.close()