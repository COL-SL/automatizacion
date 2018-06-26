from flask import request, redirect, render_template, url_for
from app import app
import locale
locale.setlocale(locale.LC_TIME, "sp") # swedish
from datetime import datetime
import time


@app.template_filter()
def format_date(date): # date = datetime object.
    t = (2009, 2, 17, 17, 3, 38, 1, 48, 0)
    t = time.mktime(t)
    return date.strftime("%b %d %Y %H:%M:%S", time.gmtime(t))

my_date=datetime.now()

month=""
if   (my_date.strftime('%m') == '01'):
    month = "Enero"
elif (my_date.strftime('%m') == '02'):
    month = "Febrero"
elif (my_date.strftime('%m') == '03'):
    month = "Marzo"
elif (my_date.strftime('%m') == '04'):
    month = "Arbil"
elif (my_date.strftime('%m') == '05'):
    month = "Mayo"
elif (my_date.strftime('%m') == '06'):
    month = "Junio"
elif (my_date.strftime('%m') == '07'):
    month = "Julio"
elif (my_date.strftime('%m') == '08'):
    month = "Agosto"
elif (my_date.strftime('%m') == '09'):
    month = "Septiembre"
elif (my_date.strftime('%m') == '10'):
    month = "Octubre"
elif (my_date.strftime('%m') == '11'):
    month = "Noviembre"
elif (my_date.strftime('%m') == '12'):
    month = "Diciembre"


day=""
if   (my_date.weekday()== 0):
    day = "Lunes"
elif (my_date.weekday() == 1):
    day = "Martes"
elif (my_date.weekday() == 3):
    day = "Miércoles"
elif (my_date.weekday() == 4):
    day = "Jueves"
elif (my_date.weekday() == 5):
    day= "Viernes"
elif (my_date.weekday() == 6):
    day = "Sábado"
elif (my_date.weekday() == 7):
    day = "Domingo"


@app.route('/')
def index():
    return render_template('index.html', name_columns=['Infinity', 'Cisco SR', 'Cisco RMA', 'Ticket SMC', 'Cliente',
                                                       'Sala de apertura', 'Adm. de circuito', 'Salas afectadas',
                                                       'País','Fecha de cierre','Escalado','Proactiva','Responsable',
                                                       'Motivo de apertura','Resolución','Tiempo abierta',
                                                       'Fecha de apertura'],
                           month_actual=month,my_date=datetime.now(),day_actual=day)

