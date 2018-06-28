from flask import request, redirect, render_template, url_for
from app import app
import locale
locale.setlocale(locale.LC_TIME, "sp") # swedish
import openpyxl
from datetime import datetime
#import numpy





#Read Excel
NAME_FILE=openpyxl.load_workbook(r'C:\Users\usr1CR\PycharmProjects\probando_jinja2\excel\Prueba.xlsx')
sheet =  NAME_FILE['Cerradas']

num_total_rows = 0
count_num_total_rows = 1
final_count_num_total_rows = 1
column_name_f= ''
next = False

while(next == False):
    column_name_f = str("f"+str(count_num_total_rows))
    #print (column_name_f)
    #print(sheet[column_name_f].value)
    if (sheet[column_name_f].value == None):
        next = True
    else:
        count_num_total_rows = count_num_total_rows + 1


for final_count_num_total_rows in range(1,count_num_total_rows):
    column_name_f = str("f" + str(final_count_num_total_rows))
    if (sheet[column_name_f].value == 'TIWS' or sheet[column_name_f].value == 'TIWS '):
        print (column_name_f)
    elif (sheet[column_name_f].value == 'TISA ' or sheet[column_name_f].value == 'TISA'):
        print(column_name_f)
    elif (sheet[column_name_f].value == 'TEDIG' or sheet[column_name_f].value == 'TEDIG '):
        print(column_name_f)

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
elif (my_date.weekday() == 2):
    day = "Miércoles"
elif (my_date.weekday() == 3):
    day = "Jueves"
elif (my_date.weekday() == 4):
    day= "Viernes"
elif (my_date.weekday() == 5):
    day = "Sábado"
elif (my_date.weekday() == 6):
    day = "Domingo"



@app.route('/')
def index():
    return render_template('index.html', name_columns=['Infinity', 'Cisco SR', 'Cisco RMA', 'Ticket SMC', 'Cliente',
                                                       'Sala de apertura', 'Adm. de circuito', 'Salas afectadas',
                                                       'País','Fecha de cierre','Escalado','Proactiva','Responsable',
                                                       'Motivo de apertura','Resolución','Tiempo abierta',
                                                       'Fecha de apertura'],
                           month_actual=month,my_date=datetime.now(),day_actual=day)

