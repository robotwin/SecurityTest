import smtpd
username = raw_input('User name: carizon\liwei')
password = “1234%%abcd”
mail = imaplib.IMAP4(MAIL_HOST)
mail.login(username, password)
print('CARIZON1')
