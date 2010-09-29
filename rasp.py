#!/usr/bin/eval PYTHONPATH=/home/krayushka/modules python
# -*- coding: utf8 -*-
import xlrd
import datetime

# settings and constants
group_sheet_name = u"Группы"
discip_sheet_name = u"Дисциплины"
rasp_sheet_name = u"Расписание"


def is_excel_date(cell):
    return cell.ctype == xlrd.XL_CELL_DATE 

def excel_date(value):
    return datetime.date(1899, 12, 31) + datetime.timedelta(value-1)  

def remove_index_from_str(s):
    ind = 0
    for i, c in enumerate(s):
        if c.isalpha():
            ind = i
            break
    return s[ind:]


def get_discip_list(sheet):
    disc = {}
    # row - disc_i
    # col - stud_i
    for stud_i in xrange(2, sheet.ncols):
        stud = sheet.cell(0, stud_i).value
        disc[ stud ] = [] 
        for disc_i in xrange(1, 58):
            sign = sheet.cell(disc_i, stud_i).value
            if ( sign != u"" ):
                wo_index = remove_index_from_str( sheet.cell(disc_i, 0).value )
                disc[stud].append( u"{0} ({1})".format(wo_index, sign) )
    
    return disc


def get_rasp(sheet):
    dates = {}
    for col in xrange(sheet.ncols):
        for row in xrange(sheet.nrows):
            current_cell = sheet.cell(row, col)
            if is_excel_date(current_cell):
                dates[ excel_date( current_cell.value) ] = (row, col)
    
    rasp = {}
    for date in dates:
        row, col = dates[date]
        cur_row = row
        rasp[date] = []
        while True:
            value_left =  sheet.cell( cur_row, col + 1 ).value
            value_right = sheet.cell( cur_row, col + 2 ).value
            if value_left != u"":
                rasp[date].append( value_left )
            if value_right != u"":
                rasp[date].append( value_right )
            cur_row += 1
            if  ((cur_row, col) in dates.values()) or (cur_row == sheet.nrows): 
                break
    return rasp


