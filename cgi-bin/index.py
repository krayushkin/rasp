#!/usr/bin/python
# -*- coding: utf8 -*

import datetime
import rasp
import sys, os
import xlrd
import codecs

#os.environ["FCGI_FORCE_CGI"] = "Y"

css_file = codecs.open("style.css", encoding="utf-8")
css_style = u""" <style type="text/css"> {0} </style> """.format(css_file.read())


def display_main():
    rasp_date_form = u"""
<form action="rasp.fcgi" method="get">
<div>
Дата: <input type="text" name="date"></input>
<input type="submit" value="Показать"></input>
</div>
</form>
"""
    return rasp_date_form


def display_rasp_by_date( rasp, date):
    next_day_link = u"""<a href="rasp.fcgi?date={0}">Вперед</a>""".format((date + datetime.timedelta(1)).strftime("%d_%m_%Y") )
    prev_day_link = u"""<a href="rasp.fcgi?date={0}">Назад</a>""".format((date - datetime.timedelta(1)).strftime("%d_%m_%Y") )
    weeks = {0 : u"Понедельник", 1 : u"Вторник", 2 : u"Среда", 3 : u"Четверг" , 4 : u"Пятница" , 5 : u"Суббота", 6 : u"Воскресенье"}


    r = None
    if rasp == []:
        r = u"Предметов в этот день нет, либо они ещё не указаны в расписании"
    else:
        unique_disc = {}
        for d in rasp:
            if d not in unique_disc:
                unique_disc[d] = 1
            else:
                unique_disc[d] += 1
        r = u"".join( [u"<li>{0} ({1})</li>".format(r, count) for r, count in unique_disc.items()] )

    return u"""<div class="nav2">{2} | {0} | {3}</div><div class="weekday">{4}:</div>
   <ul>{1}</ul>
    """.format(date, r, prev_day_link, next_day_link, weeks[date.weekday()]) 

def display_all_disc(student, stud_disc):
    d = u"".join( [u"<li>{0}</li>".format(d) for d in sorted(stud_disc)] )
    return u"""
    <div class="nav2">Дисциплины</div>    
    <div class="weekday">{0}</div>
    <ol>{1}</ol>
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
    start_response('200 OK', [('Content-Type', 'text/html; charset=UTF-8') ])
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
    <meta http-equiv="Content-Type" content="application/xhtml+xml; charset=utf-8" />
    <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" /> 
    <meta name="MobileOptimized" content="240" /> 
    <meta name="PalmComputingPlatform" content="true" /> 
    <title>Расписание</title>
     {0}
  </head>
<body>
    """.format(css_style)

    today_link = u"""<a href="rasp.fcgi?date=today">Сегодня</a>"""
    tomorrow_link = u"""<a href="rasp.fcgi?date=tomorrow">Завтра</a>"""
    
    if len(form) == 0:
        form["date"] = "today"

    if "date" in form:
        if form["date"] == "today":
            today_link = u"Сегодня"
        elif form["date"] == "tomorrow":
            tomorrow_link = u"Завтра"

        elif to_date( form["date"] ) == datetime.date.today():
            today_link = u"Сегодня"
        elif to_date( form["date"] ) == (datetime.date.today() + datetime.timedelta(1)):
            tomorrow_link = u"Завтра"

    yield u"""<div class="nav1">{0}  |  {1}</div>""".format(today_link, tomorrow_link)

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
    students_list = sorted(name_discip_dict.keys())
    first_stud = students_list.pop(0)
    first_stud_opt = u""" <option selected="selected">{0}</option>""".format(first_stud)
    options_stud = first_stud_opt + u"". join ([ u"<option>{0}</option>".format(stud) for stud in students_list] )
    disc_form = u"""
<div class="discform">
  <form action="rasp.fcgi" method="get">
    <div>
      <select name="stud">{0}</select>
      <input type="submit" value="Дисциплины" />
    </div>
  </form>
</div>
""".format(options_stud)
    yield disc_form
    yield u"""
<div class="foot">
  <div>
    Расписание ЦОО ФИСТ ЭВМ alpha {0}
  </div>
    Заметили ошибку или есть пожелания? <a href="mailto:cr0ss@mail.ru">Пишите</a>
</div>
""".format(escape(u"© 2010 Краюшкин Дмитрий"))
    yield u"""
</body>
</html>"""

WSGIServer(app).run()

