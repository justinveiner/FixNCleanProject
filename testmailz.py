import smtplib

sender = input("\nPlease enter your Gmail email address\n->")
username = sender
password = input("\nPlease enter the password to your Gmail account (this information will not be permanently saved/stored)\n->")

receivers = 'jav14@me.com'

message = """From: Fix 'N' Clean Team 'justin.veiner@gmail.com'
To: To Person 16jav1@queensu.ca
Subject: SMTP e-mail test

This is a test e-mail message.
"""
#AUTHENTICATE
try:
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(username,password)
    smtpObj.sendmail(sender, receivers, message)
    print("Successfully sent email")
    smtpObj.close()
except:
    print("Error: unable to send email")
##print("Error: unable to send email")
#mail.protection.outlook.com
#smtpObj = smtplib.SMTP('mail.protection.outlook.com', 587)

# ENABLE LESS SECURE APPS IN GMAIL