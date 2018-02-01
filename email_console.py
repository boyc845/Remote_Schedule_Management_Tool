import smtplib
import datetime
from email.mime.text import MIMEText
from email.header import Header
from collections import namedtuple

import date_reminder
import file_console

EmailConsole = namedtuple('EmailConsole', ['email_smtp', 'sender', 'receiver', 'excel'])


DEBUG = True
##163   smpt.163.com:465/994   pop.163.com:995    imap.163.com:993
##content = 'example'
##
##gmail = smtplib.SMTP('smtp.gmail.com', 587)
##
##gmail.ehlo()
##
##gmail.starttls()
##
##gmail.login('xiguil@uci.edu', '.fly84597245.')
##
##gmail.sendmail('xiguil@uci.edu', 'boyc845@gmail.com', content)
##
##gmail.close()

#console#######################################################
def conncet_SMTP_server(server: str, port: int):
    '''connect to smtp server'''
    mail_smtp = smtplib.SMTP(server, port)
    mail_smtp.ehlo()
    mail_smtp.starttls()
    return mail_smtp

def login_email_account(mail_smtp: smtplib.SMTP, email_adress: str, password: str):
    '''login in email account'''
    mail_smtp.login(email_adress, password)

def send_mail(mail_smtp: smtplib.SMTP, sender: str, receiver: str, content):
    '''send an email'''
    mail_smtp.sendmail(sender, receiver, content)

def get_content(text: str, subject: str) ->str:
    '''ask user to enter content that he wants to send'''
    msg = MIMEText(text)
    msg['Subject'] = Header('**自动提醒**文件: '+ subject)
    return msg

def save_as_default(emailconsole: EmailConsole, path: str):
    '''save login information in user_data.txt  as default setting'''
    outfile = open('user_data.txt', 'w')
    for index in emailconsole:
        outfile.write(index)
    outfile.write(path)

def automatic_login(list_data: list):
    '''automatic login email'''
    excel = list_data[0]
    e_smtp = conncet_SMTP_server(list_data[1], int(list_data[2]))
    login_email_account(e_smtp, list_data[3], list_data[4])
    receiver_email = list_data[5]
    return EmailConsole(email_smtp = e_smtp,
                        sender = list_data[3],
                        receiver = receiver_email,
                        excel = excel)

def get_date_and_send(emailconsole: EmailConsole):
    '''get date from excel and send email'''
    er = date_reminder.excel_reader(emailconsole.excel)
    text = er.create_message()
    message = get_content(text, er.get_file_name())
    
    send_mail(emailconsole.email_smtp,
              emailconsole.sender,
              emailconsole.receiver, message.as_string())
    emailconsole.email_smtp.close()
    
if __name__ == '__main__':

##        time.sleep(date_reminder.delaytime())
    emailconsole = automatic_login(file_console.read_data('user_data.txt'))
    get_date_and_send(emailconsole)
