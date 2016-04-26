from flask import Flask, render_template

import os

import json

from pwmodel import PwModel, pw_model_from_excel

app = Flask(__name__)

APP_ROOT = os.path.dirname(os.path.abspath(__file__))   # refers to application_top
MODEL_ROOT = APP_ROOT
APP_STATIC = os.path.join(APP_ROOT, 'static')

@app.route('/')
def hello_world():
    return render_template('index.html', username = 'Elena')

@app.route('/<uname>')
def hello_world2(uname):
    allData = get_all_data(uname, uname)
    jsonAllData = json.dumps(allData)
    return render_template('index.html', allData = jsonAllData)


def get_all_data(uname, model):
    md = pw_model_from_excel(os.path.join(MODEL_ROOT, "Jobs.xlsx"))
    return(md.getAllCalcs())

if __name__ == '__main__':
    app.run()
