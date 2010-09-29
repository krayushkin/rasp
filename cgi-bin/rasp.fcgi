#!/usr/bin/eval PYTHONPATH=/home/krayushka/modules python
# -*- coding: utf8 -*


import datetime
import rasp
import sys, os
import xlrd



def display_main():
    rasp_date_form = u"""<FORM action="rasp.fcgi" method="get">
Дата: <INPUT type="text" name="date">
<INPUT type="submit" value="Показать">
 </FORM>"""
    return rasp_date_form


def display_rasp_by_date( rasp, date):
    next_day_link = u"""<a href="rasp.fcgi?date={0}">Вперед</a>""".format((date + datetime.timedelta(1)).strftime("%d_%m_%Y") )
    prev_day_link = u"""<a href="rasp.fcgi?date={0}">Назад</a>""".format((date - datetime.timedelta(1)).strftime("%d_%m_%Y") )
    weeks = {0 : u"Понедельник", 1 : u"Вторник", 2 : u"Среда", 3 : u"Четверг" , 4 : u"Пятница" , 5 : u"Суббота", 6 : u"Воскресенье"}


    r = None
    if rasp == []:
        r = u"Предметов в этот день нет, либо они не указаны в расписании<br>"
    else:
        unique_disc = {}
        for d in rasp:
            if d not in unique_disc:
                unique_disc[d] = 1
            else:
                unique_disc[d] += 1
        r = u"".join( [u"{0} ({1})<br>".format(r, count) for r, count in unique_disc.items()] )

    return u"""<p>{2} | {0} | {3}<br><br>{4}</p>
    {1}<br>
    """.format(date, r, prev_day_link, next_day_link, weeks[date.weekday()]) 

def display_all_disc(student, stud_disc):
    d = u"".join( [u"<li>{0}</li>".format(d) for d in sorted(stud_disc)] )
    return u"""
    <p><b>Дисциплины</b></p>
    <p>{0}</p>
    <ol>{1}</ol><br>
    """.format(student, d)

def to_date(string_date):
    l = string_date.split("_")
    if len(l) == 3:
        try:
            day = int(l[0])
            month = int(l[1])
            year = int(l[2])
        except ValueError:
            raise ValueError(u"Неправильный формат даты".encode("utf-8"))
        return datetime.date(year, month, day)
    else:
        raise ValueError(u"Неправильный формат даты. В дате должно быть 3 компоненты".encode("utf-8"))


# -----------------------------------------------------------------------------


from cgi import escape
from flup.server.fcgi import WSGIServer
import urlparse

book = xlrd.open_workbook("rasp.xls")
name_discip_dict = rasp.get_discip_list(book.sheet_by_index(1) )
date_rasp_dict = rasp.get_rasp(book.sheet_by_index(0) )

def app(environ, start_response):
    start_response('200 OK', [('Content-Type', 'text/html'), ('Content-Encoding', 'utf-8') ])
    for s in app_helper(environ, start_response):
        yield s.encode("utf-8")

def app_helper(environ, start_response):
    qs = urlparse.parse_qs(environ["QUERY_STRING"])
    form = {}
    for k, v in qs.items():
        form[k] = v[0]
    yield u"""
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
      "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

    <html xmlns="http://www.w3.org/1999/xhtml">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        
        <title>Расписание</title>
      </head>
      <body>
    """

    today_link = u"""<a href="rasp.fcgi?date=today">Сегодня</a>"""
    tomorrow_link = u"""<a href="rasp.fcgi?date=tomorrow">Завтра</a>"""
    
    if len(form) == 0:
        form["date"] = "today"

    if "date" in form:
        if form["date"] == "today":
            today_link = u"<b>Сегодня</b>"
        elif form["date"] == "tomorrow":
            tomorrow_link = u"<b>Завтра</b>"

        elif to_date( form["date"] ) == datetime.date.today():
            today_link = u"<b>Сегодня</b>"
        elif to_date( form["date"] ) == (datetime.date.today() + datetime.timedelta(1)):
            tomorrow_link = u"<b>Завтра</b>"

    yield u"{0}  |  {1}<br>".format(today_link, tomorrow_link)

    if "date" in form:
        date = None
        if form["date"] == "today":
            
            date = datetime.date.today()
        elif form["date"] == "tomorrow":
            date = datetime.date.today() + datetime.timedelta(1)
        else:
            date = to_date( form["date"] )

        if date in date_rasp_dict:
            yield display_rasp_by_date( date_rasp_dict[date], date )
        else:
            yield display_rasp_by_date( [], date )
        
    elif "stud" in form :
        stud_name =  form["stud"].decode("utf-8") 
        if stud_name in name_discip_dict: 
            yield display_all_disc(stud_name, name_discip_dict[ stud_name ] )
        else:
            yield u"<p>Такого студента нет в базе</p>"

    options_stud = u"". join ([ u"<option>{0}</option>".format(stud) for stud in sorted(name_discip_dict.keys())] )
    disc_form = u"""  
    <FORM action="rasp.fcgi" method="get">
     <br>Дисциплины: <SELECT name="stud">{0}</SELECT><INPUT type="submit" value="Показать"></FORM>""".format(options_stud)
    yield disc_form
    yield u"""
<p>Расписание ЦОО ФИСТ ЭВМ alpha {0}</p>
Заметили ошибку или есть пожелания? <a href="mailto:cr0ss@mail.ru">Пишите</a>
""".format(escape(u"© 2010 Краюшкин Дмитрий"))
    yield u"""
    </body>
    </html>"""

WSGIServer(app).run()

