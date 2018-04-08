from flask import request, redirect, render_template, url_for
from app import app
from datetime import datetime
import time

@app.template_filter()
def format_date(date): # date = datetime object.
    t = (2009, 2, 17, 17, 3, 38, 1, 48, 0)
    t = time.mktime(t)
    return date.strftime("%b %d %Y %H:%M:%S", time.gmtime(t))

@app.route('/')
def index():
    return render_template('index.html', cosas=['RPG', 'Python', 'Juegos de mesa', 'Cthulhu', 'etc'], my_date=datetime.now())