#encoding: utf-8

from datetime import datetime

import xlwt
import xlrd

USERS = [
    {'name': u'Tony Stark', 'email': u'tony@stark.com'},
    {'name': u'Steve Rogers', 'email': u'rogers@usa.com'},
    {'name': u'Bruce Banner', 'email': u'bruce@university.com'},
]

BANK = [
    {
        'name': u'Tony Stark', 
        'balance': 100300.50, 
        'updated_on': datetime(year=2013, month=1, day=10, 
            hour=23, minute=10, second=13),
    },
    {
        'name': u'Steve Rogers', 
        'balance': 504.90, 
        'updated_on': datetime(year=1939, month=5, day=25,
            hour=13, minute=14, second=40),
    },
    {
        'name': u'Bruce Banner', 
        'balance': 4301,
        'updated_on': datetime(year=2010, month=3, day=13,
            hour=9, minute=26, second=33),
    },
]

def simple_workbook():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(u'Users')
    worksheet.write(0, 0, u'Name')
    worksheet.write(0, 1, u'E-mail')
    x = 0
    for user in USERS:
        x = x + 1
        worksheet.write(x, 0, user['name'])
        worksheet.write(x, 1, user['email'])
    return workbook

def format_numbers_workbook():
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(u'Users')
    worksheet.write(0, 0, u'Name')
    worksheet.write(0, 1, u'Balance')
    worksheet.write(0, 2, u'Updated on')
    x = 0
    for account in BANK:
        x = x + 1
        worksheet.write(x, 0, account['name'])
        worksheet.write(x, 1, account['balance'],
            style=xlwt.easyxf(num_format_str="#,###.00"),
        )
        worksheet.write(x, 2, account['updated_on'], 
            style=xlwt.easyxf(num_format_str='dd/mm/yyyy hh:mm:ss'),
        )
    return workbook

def workbook_with_borders():
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.MEDIUM
    borders.right = xlwt.Borders.MEDIUM
    borders.top = xlwt.Borders.MEDIUM
    borders.bottom = xlwt.Borders.MEDIUM
    style = xlwt.XFStyle()
    style.borders = borders
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(u'Users')
    worksheet.write(1, 0, u'Name', style=style)
    worksheet.write(1, 1, u'E-mail', style=style)
    x = 1
    for user in USERS:
        x = x + 1
        worksheet.write(x, 0, user['name'], style=style)
        worksheet.write(x, 1, user['email'], style=style)
    return workbook

if __name__ == "__main__"    :
    workbook = simple_workbook()
    workbook.save('users.xls')
    workbook = workbook_with_borders()
    workbook.save('users_borders.xls')
    workbook = format_numbers_workbook()
    workbook.save('bank.xls')