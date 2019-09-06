import os, sys

path = os.path.dirname(os.path.abspath(__file__))
template_path = path + '/templates'
paths = sys.path
if path not in paths:
    sys.path.append(path)
    sys.path.append(template_path)

import flask
from flask import Flask
from flask_wtf import Form
from wtforms import TextField
import forms
import ss_gen
from flask import Flask, Response, redirect, url_for, request, \
    session, abort, render_template, flash
from flask_login import LoginManager, UserMixin, login_required, \
    login_user, logout_user


app = Flask(__name__)
app.secret_key = 'development key'
login = LoginManager(app)

@app.route('/ss_generator', methods = ['GET', 'POST'])
def ss_generator():
    form = forms.SSGeneratorForm()
    if request.method == 'POST':
        web_data_arr = [form.emp_id.data, form.ss_id.data, form.emp_db_id.data, form.ws_name.data,
                        True, form.all_colleagues.data, form.email.data, form.pwd.data]
        ss_gen.generate(web_data_arr)
        return Response(ss_gen.generate(web_data_arr), mimetype='text/html')
    elif request.method == 'GET':
        return render_template('ss_generator.html', form = form)

if __name__ == '__main__':
    app.run(debug = True)
