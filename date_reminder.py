import xlrd
import datetime
from collections import namedtuple
from collections import OrderedDict

DEBUG = True

Item = namedtuple('Item', ['item_number', 'customer', 'merchandise', 'quantity', 'delivery_time', 'production_location', 'date_content'])

def get_today_date() -> 'date':
    '''get current date'''
    now = datetime.datetime.now()
    return now.date()

def search_specific_word(word: str, sheet: 'sheet') -> list:
    '''search all cells which contain the word in a sheet and then return the list of position of the word's cell'''
    list_related_cell_positions = []
    for rownum in range(sheet.nrows):
        for colnum in range(sheet.ncols):
            if word in str(sheet.cell_value(rownum, colnum)):
                list_related_cell_positions.append((rownum, colnum))
    return list_related_cell_positions
#----------------------------------------------------------get information for namedtuple
def get_item_number(item_row: int, sheet: 'sheet') -> str:
    '''return style number'''
    for row, col in search_specific_word('款号', sheet):
        return str(sheet.cell_value(item_row, col))

def get_customer(item_row: int, sheet: 'sheet') -> 'date':
    '''return customer name'''
    for row, col in search_specific_word('客人', sheet):
        return sheet.cell_value(item_row, col)

def get_merchandise(item_row: int, sheet: 'sheet') -> str:
    '''return merchadise's name'''
    for row, col in search_specific_word('跟单', sheet):
        return sheet.cell_value(item_row, col)

def get_delivery_time(item_row: int, sheet: 'sheet') -> 'date':
    '''return delivery time'''
    for row, col in search_specific_word('货期', sheet):
        if sheet.cell_type(item_row, col) == 3:
            time_tuple = xlrd.xldate_as_tuple(sheet.cell_value(item_row, col), 0)
            return datetime.date(time_tuple[0], time_tuple[1],time_tuple[2])
    return 0

def get_quantity(item_row: int, sheet: 'sheet') -> 'int':
    '''return quantity'''
    try:
        for row, col in search_specific_word('数量', sheet):
            return int(sheet.cell_value(item_row, col))
    except:
        return 0

def get_production_loaction(item_row: int, sheet: 'sheet') -> str:
    '''return production location'''
    for row, col in search_specific_word('组别', sheet):
        return sheet.cell_value(item_row, col)


def create_item_namedtuple(item_row: int, sheet: 'sheet') -> Item:
    '''create item namedtuple'''
    return Item(item_number = str(get_item_number(item_row, sheet)),
                customer = str(get_customer(item_row, sheet)),
                merchandise = str(get_merchandise(item_row, sheet)),
                delivery_time = get_delivery_time(item_row, sheet),
                date_content = get_content_row(item_row, sheet),
                quantity = get_quantity(item_row, sheet),
                production_location = get_production_loaction(item_row, sheet))
#----------------------------------------------------------------------------------

def find_superior_label(date_label_position: '(row,col)', sheet: 'sheet') -> str:
    ''' '''
    row, col = date_label_position
    if row-1 < 0:
        return ''
    for index in range(col):
        if sheet.cell_value(row-1, col- index) != '':
            return sheet.cell_value(row-1, col- index)
    return ''

def get_list_superior_label(sheet: 'sheet') -> list:
    ''' return a list of superior in order for a sheet'''
    result = []
    date_label_positions = search_specific_word('日期', sheet)
    for row, col in date_label_positions:
        if find_superior_label((row, col), sheet) not in result:
            result.append(find_superior_label((row, col), sheet))
    return result

def get_content_row(item_row: int, sheet: 'sheet') -> 'OrderedDict':
    '''get all date in a row'''
    date_label_positions = search_specific_word('日期', sheet)
    content = OrderedDict()
    for y, x in date_label_positions:
        if sheet.cell_type(item_row, x) == 3:
            if find_superior_label((y,x), sheet) not in content:
                content[find_superior_label((y,x), sheet)] = []
            time_tuple = xlrd.xldate_as_tuple(sheet.cell_value(item_row, x), 0)
            content[find_superior_label((y,x), sheet)].append((sheet.cell_value(y, x),
                                                               datetime.date(time_tuple[0],
                                                                             time_tuple[1],
                                                                             time_tuple[2])))
    return content

def delaytime() -> int:
    ''' '''
    next_day = datetime.datetime.combine(get_today_date()+datetime.timedelta(1),datetime.time(8,0,0))
    print(next_day)
    sec = next_day - datetime.datetime.now()
    return sec.total_seconds()

#------------------------create message
def for_list_message_date_to_reminder(list_date: list, itemnumber: str, deliverydate: 'date') -> str:
    '''create message by date list'''
    count = 0
    message = ''
    for label, date in list_date:
        count += 1
        word = label.strip('日期')
        if date >= get_today_date():
            message += ('\n{}阶段，未超时\n{}货期{}，{}未完成，截止时间{}，\
请按时寄出，不然会耽误货期.\n').format(word, itemnumber, deliverydate, word, date)
        else:
            message += ('\n{}阶段，超时\n{}货期{}，{}未完成，截止时间{}，\
已经超时{}天，请抓紧时间不然会耽误货期.\n').format(word, itemnumber, deliverydate, word, date,
                                 (get_today_date()- date).days)
        if count == 3:
            return message
    return message

def for_content_message_date_to_reminder(datecontent: dict, itemnumber: str, deliverydate: 'date') -> str:
    '''create message by date content'''
    message = ''
    count = 0
    for key in datecontent.keys():
        count += 1
        message += '\n' + key
        message += for_list_message_date_to_reminder(datecontent[key], itemnumber, deliverydate)
        if count == 3:
            return message
    return message
            

def for_item_message_date_to_reminder(item: Item) -> str:
    '''create message by item '''
    message = ''
    count = 0
    if item.delivery_time == 0 or item.date_content == {}:
        return message
    message = '\n\n提醒\n客户: {} 款号: {} 数量: {} 货期: {} 组别: {}'.format(item.customer, item.item_number,
                                                                  item.quantity, item.delivery_time,
                                                                  item.production_location)
                                                                  
    message += for_content_message_date_to_reminder(item.date_content, item.item_number, item.delivery_time)
    return message

def for_sheet_message_date_to_reminder(list_items: list) -> str:
    ''' create message by sheet'''
    message = ''
    for item in list_items:
        message +=for_item_message_date_to_reminder(item)
    return message

def for_book_message_date_to_reminder(book: 'OrderedDict'):
    message = ''
    for key in book.keys():
        message += ('\n------------{}------------\n{}\n').format(key,for_sheet_message_date_to_reminder(book[key]))
    return message
#------------------------------------------------------------------------------------------

def get_date_onesheet(sheet: 'sheet') -> list:
    '''get date base on_label col'''
    list_items = []
    date_label_positions = search_specific_word('日期', sheet)
    
    if date_label_positions == []:
        return list_items
    start_rownum = date_label_positions[0][0]
    
    for rownum in range(start_rownum + 1, sheet.nrows):
        list_items.append(create_item_namedtuple(rownum, sheet))
    return list_items

def get_date_wholebook(sheets):
    '''get date whole book '''
    all_stuff = OrderedDict()
    for sheet in sheets:
        all_stuff[sheet.name] = get_date_onesheet(sheet)
    return all_stuff

###interface
def get_file_path() -> str:
    ''' '''
    path = input('输入文件路径（包括文件后缀，例如：.xlsx）: ')
    return path

def read_file(path: str) -> 'class excel':
    excel = excel_reader(path)
    return excel

    
    

class excel_reader:
    def __init__(self, file_name: str):
        '''try to open an excel file. raise an error, if can't'''

        self.excel_file = xlrd.open_workbook(file_name)
        self.file_name = file_name
        self.sheets = self.excel_file.sheets()
        self.data_label_positions_allsheet = []
        self.all_stuff = get_date_wholebook(self.sheets)

    def get_file_name(self) ->str:
        '''return excel file name'''
        return self.file_name

    def return_all(self):
        ''' '''
        return self.all_stuff

    def create_message(self):
        ''' '''
        message = for_book_message_date_to_reminder(self.all_stuff)
        return message
        
    def try_function(self):
        return get_content_row(2, self.sheets[0])



if __name__ == '__main__':
    while True:
        try:
            path = get_file_path()
            break
        except:
            print('没有该文件')
    er = excel_reader(path)
    alll = er.return_all()
    message = for_book_message_date_to_reminder(alll)
##    sheet1 = excel.return_all()['跟踪提醒表']
