import poplib
import email
import re
from email import header
from collections import OrderedDict
from collections import namedtuple
import datetime

EmailContent = namedtuple('emailcontent', ['filename', 'receive_date', 'content_command'])

CONFIG = 'receive_time.txt'

def getYear(date):
    rslt = re.search(r'\b2\d{3}\b', date)
    return int(rslt.group())

def getMonth(date):
    monthMap = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12,}

    rslt = re.findall(r'\b\w{3}\b', date)
    for i in range(len(rslt)):
        month = monthMap.get(rslt[i])
        if None != month:
            break

    return month

def getDay(date):
    rslt = re.search(r'\b\d{1,2}\b', date)
    return int(rslt.group())

def getTime(date):
    rslt = re.search(r'\b\d{2}:\d{2}:\d{2}\b', date)
    timeList = rslt.group().split(':')

    for i in range(len(timeList)):
        timeList[i] = int(timeList[i])

    return timeList

def transformDate(date):
    '''transform date from mail date'''
    list_time = getTime(date)
    return datetime.datetime(getYear(date),
                         getMonth(date),
                         getDay(date), list_time[0], list_time[1],list_time[2])

def get_today_date() -> 'date':
    '''get current date'''
    return datetime.date.today()


def getRecentReadMailTime():
    fp = open(CONFIG, 'r')
    rrTime = fp.read()
    fp.close()
    return rrTime

def setRecentReadMailTime():
    fp = open(CONFIG, 'w')
    fp.write(datetime.datetime.now().ctime())
    fp.close()
    return

def parseMailSubject(msg) -> str:
    '''parse mail subject and return subject'''
    subStr = msg.get('subject')
    if subStr == None:
        subject = 'No Subject'
    else:
        subList = header.decode_header(subStr)
        subinfo = subList[0][0]
        subcode = subList[0][1]

        if isinstance(subinfo, bytes):
            subject = subinfo.decode(subcode)
        else:
            subject = subinfo
    return subject

def parseMailContent(msg) -> str:
    '''parse mail content and return content'''
    listmsg = []
    if msg.is_multipart():
        for part in msg.get_payload():
            listmsg.extend(parseMailContent(part))
    else:
        bMsgStr = msg.get_payload(decode=True)
        charset = msg.get_param('charset')
        msgStr = 'Decode Failed'
        try:
            if charset == None:
                msgStr = bMsgStr.decode()
            else:
                msgStr = bMsgStr.decode(charset)
        except:
            print('decode failed')
            pass
        listmsg.append(msgStr)
    return listmsg
#create filename,command,receive date--------------------------------------------
def get_filename_from_subject(subject: str):
    '''get filename from email subject'''
    filename = ''
    if '**自动提醒**' in subject:
        filename = subject[(subject.find('件')+3):]
    return filename

def get_command_from_message(msg: str) ->dict:
    '''get command from email'''
    dict_command = {}
    list_msg = msg.split('\r\n')
    for msg in list_msg:
        if msg.startswith('完成'):
            list_command = msg.split()
            for command in list_command[2:]:
                if list_command[1] not in dict_command:
                    dict_command[list_command[1]] = [command]
                else:
                    dict_command[list_command[1]].append(command)
    return dict_command

def get_command_from_mail(mail: list) -> list:
    '''create command from emails'''
    list_dict_command = []
    for msg in mail:
        list_dict_command.append(get_command_from_message(msg))
    return list_dict_command

def create_emailcontent_namedtuple(emailmsg) -> 'emailcontent':
    '''create email content namedtuple'''
    return EmailContent(filename = get_filename_from_subject(parseMailSubject(emailmsg)),
                        receive_date = transformDate(emailmsg.get('Date')),
                        content_command = get_command_from_mail(parseMailContent(emailmsg)))

def create_list_emailcontent(emailmsgs:list) -> list:
    '''create list of email content namedtuple'''
    lis = []
    for emailmsg in emailmsgs:
        lis.append(create_emailcontent_namedtuple(emailmsg))
    return lis
#---------------------------------------------------------

def get_emailmsg(emailserver: 'email'):
    '''return emailmsg '''
    list_emails = []
    mailCount, size = emailserver.stat()
    mailNumList = list(range(mailCount))
    mailNumList.reverse()
    parsedMsg = OrderedDict()
    hisTime = transformDate(getRecentReadMailTime())
    for index in mailNumList:
        message = emailserver.retr(index+1)[1]
        mail = email.message_from_bytes(b'\n'.join(message))
        if transformDate(mail.get('Date')) > hisTime:
            list_emails.append(mail)
    return list_emails

class EmailReader:
    def __init__(self, popServer: str, username: str, password: str):
        self.emailServer = poplib.POP3(popServer)
        self.emailServer.user(username)
        self.emailServer.pass_(password)

    def get_mails_object(self):
        '''return parsed message'''
        return get_emailmsg(self.emailServer)
    
    def return_emailserver(self):
        return self.emailServer


if __name__ == '__main__':
    pass

    print('login in...')
    emailreader = EmailReader('pop.263.net', 'reminder@sharicca.com',
                              'kingrik845')
    print('success')
    print('reading email...')
    Msg = emailreader.get_mails_object()
    print('done')
    lis = create_list_emailcontent(Msg)










        
