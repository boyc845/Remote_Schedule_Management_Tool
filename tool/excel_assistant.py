#
import datetime
import logging
from threading import Thread
import time

import file_console
import date_reminder
import email_console
import excel_editer
import email_receiver

logging.basicConfig(filename = 'logfile.log', level = logging.DEBUG)

#sending email-----------------------------------------------------------------------------------
def get_server_host() -> str:
    '''host'''
    host = input('服务器: ').strip()
    return host

def get_port() -> int:
    '''port'''
    port = int(input('端口: ').strip())
    return port

def get_email_adress() -> str:
    ''' ask user to enter an email adress'''
    while True:
        email_adress = input('输入邮箱账号: ').strip()
        if '@' in email_adress:
            return email_adress
        else:
            print('请输入正确的邮箱地址')

def get_password() -> str:
    ''' ask user to enter a password'''
    password = input('输入邮箱密码: ').strip()
    return password

def get_receiver() -> str:
    '''ask user to enter receiver email'''
    print('\n提醒人邮箱')
    receiver = get_email_adress().strip()
    return receiver

def send_email(emailconsole: 'EmailConsole'):
    '''send email'''
    while True:
        try:
            print('收件人: ' + emailconsole.receiver)
            print('发送中...')
            email_console.get_date_and_send(emailconsole)
            print('已发送\n时间: {}'.format(datetime.datetime.now()))
            break
        except Exception as detail:
            logging.debug('{} {} {}'.format('automatic_login', datetime.datetime.now(), detail))
            print('发送失败,60秒后重新发送')
            time.sleep(60)
            emailconsole = default_login()

def get_login_info() -> email_console.EmailConsole:
    '''login interface. return email_console.EmailConsole'''
    while True:
        path = date_reminder.get_file_path()
        if file_console.file_whether_exist(path):
            excel = path
            break
        else:
            print('没有该文件')
            
    while True:
        try:
            host = get_server_host()
            port = get_port()
            print('连接中...')
            e_smtp = email_console.conncet_SMTP_server(host, port)
            print('连接成功')
            break
        except Exception as detail:
            logging.debug('{} {} {}'.format('connect serve', datetime.datetime.now(), detail))
            print('无法连接，服务器或端口错误.')

    while True:
        try:
            user_email = get_email_adress()
            password = get_password()
            print('账号: ' + user_email)
            print('登录中...')
            email_console.login_email_account(e_smtp, user_email, password)
            print('登录成功')
            break
        except Exception as detail:
            logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
            print('无法登录，账号或密码错误')

    receiver_email = get_receiver()

    file_console.save_data('user_data.txt', [path, host, str(port),user_email, password,
                                             receiver_email])
                                             
    return email_console.EmailConsole(email_smtp = e_smtp,
                        sender = user_email,
                        receiver = receiver_email,
                        excel = excel)

def default_login():
    '''login with default login information'''
    if file_console.file_whether_exist('user_data.txt'):
        print('自动登录中...')
        emailconsole = email_console.automatic_login(file_console.read_data('user_data.txt'))
        print('登录成功')
        return emailconsole
    else:
        print('无法通过记录登陆，请重新输入')
        emailconsole = get_login_info()
        return emailconsole

#write data----------------------------------------------------------------------------------------
def auto_write_data(excel_filename: str):
    '''auto write data to excel'''
    if not file_console.file_whether_exist(excel_filename):
        print('文件不存在')
        return
    excelediter = excel_editer.Excel_editer(excel_filename)
    try:
        excelediter.write_date_to_rows()
        print('自动写入成功')
    except Exception as detail:
        logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
        print('无法写入.')
    try:
        excelediter.save_excelfile()
    except Exception as detail:
        logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
        print('请先关闭准备写入的excel文件')

#receive email------------------------------------------------------------------------------------
def receive_email_login() -> 'EmailReader':
    '''login and receive email'''
    if file_console.file_whether_exist('user_data.txt'):
        list_data = file_console.read_data('user_data.txt')
        #####pop adress
        try:
            emailreader = email_receiver.EmailReader('pop.263.net', list_data[3], list_data[4])
        except Exception as detail:
            logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
            print('接收邮件登录失败')
    return emailreader

def check_mailbox() -> str:
    '''receive email and return message'''
    print('等待接收新邮件...')
    while True:
        time.sleep(60)
        emailreader = receive_email_login()
        try:
            Msg = emailreader.get_mails_object()
            if Msg == []:
                continue
            list_emailcontent = email_receiver.create_list_emailcontent(Msg)
            print('收到指令,准备写入')
            for emailcontent in list_emailcontent:
                if emailcontent.content_command == []:
                    continue
                print(emailcontent.filename)
                exceler = excel_editer.Excel_editer(emailcontent.filename)
                for commands in emailcontent.content_command:
                    exceler.write_done(commands, emailcontent.receive_date.date())
                exceler.save_excelfile()
                print(emailcontent.filename + '写入成功,已保存.')
                send_email(es)
            email_receiver.setRecentReadMailTime()
            print('等待接收新邮件...')
            break
        except Exception as detail:
            print('收邮件失败或写入失败,请关闭文件，60秒后自动重试')
            logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
            continue
                
##                for key in Msg.keys():
##                    if excel_editer.get_filename_from_subject(key) != '':
##                        print('收到指令,准备写入')
##                        filename = excel_editer.get_filename_from_subject(key)
##                        print(filename)
##                        dict_command = excel_editer.get_command_from_message(Msg[key][0])
##                        if dict_command == {}:
##                            print('收到指令错误byebye')
##                            continue
##                        ###############
##                        receive_date = email_receive
##                        exceler = excel_editer.Excel_editer(filename)
##                        exceler.write_done(dict_command)
##                        exceler.save_excelfile()
##                        print('写入成功,已保存.')
##                        es = default_login()
##                        send_email(es)
##                        email_receiver.setRecentReadMailTime()                    
##                print('等待接收新邮件...')
##        except Exception as detail:
##            print('收邮件失败或写入失败,请关闭文件，60秒后自动重试')
##            logging.debug('{} {} {}'.format('login account', datetime.datetime.now(), detail))
##            continue

def menu(emailconsole):
    ''' menu'''
    print('菜单')
    print('a.自动写入数据\nb.默认修改登录信息\nc.结束菜单\n')
    while True:
        command = input('输入字母: ').strip()
        if command == 'a':
            print('1.自动写入数据')
            filename = input('输入要写入文件的路径, 按回车为默认文件: ')
            if filename == '':
                auto_write_data(emailconsole.excel.get_file_name())
                break
            else:
                auto_write_data(filename)
                break
        elif command == 'b':
            get_login_info()
        elif command == 'c':
            break
        else:
            print('请输入正确字母')
        
#time.Thread--------------------------------------------------------------------------------------
def timer_auto_send():
    '''auto send email'''
    while True:
        print('下次发送时间: ')
        time.sleep(date_reminder.delaytime())
        send_email(default_login())

if __name__ == '__main__':
    es = default_login()
    auto_write_data(es.excel)
    send_email(es)
    t1 = Thread(target = timer_auto_send)
    t2 = Thread(target = check_mailbox)
    t1.start()
    time.sleep(5)
    t2.start()
