import smtplib

try:
		msgFrom = str(input("Informe o e-mail de destino: "))
		smtpObj = smtplib.SMTP('smtp.outlook.com', 587)
		smtpObj.ehlo()
		smtpObj.starttls()
		msgTo = 'Email de envio'
		toPass = 'Senha'
		smtpObj.login(msgTo, toPass)
		msg = '''
		Mensagem do E-mail
		'''
		smtpObj.sendmail(msgTo,msgFrom,'Subject: Titulo do email\n{}'.format(msg))
		smtpObj.quit()
		print("Email enviado com sucesso!")
except:
		print("Erro ao enviar e-mail")

'''
Linha 1: Importamos a biblioteca smtplib padrão do python;
Linha 4: Armazenamos o e-mail de destino na variável msgFrom;
Linha 5: Iniciamos uma conexão smtp com o servidor do Outlook na porta especificada;
Linha 6: Depois de termos um objeto SMTP devemos chamar a função ehlo() para dar um “Hello” ao servidor do Outlook ao estabelecermos a conexão;
Linha 7: Como estamos utilizando a porta 587, precisaremos chamar a função starttls(), que irá criptografar nossa conexão;
Linha 10: A função login(email,password) irá logar nossas credenciais no servidor;
Linha 12: sendmail(to,from,msg) É a função que envia e-mail passando como argumentos quem está enviando, para quem e a mensagem;
Linha 13: E por fim devemos fechar a conexão com a função quit();
'''