from flask_wtf import Form
from wtforms import TextField, IntegerField, TextAreaField, SubmitField, RadioField, SelectField, BooleanField, PasswordField
from wtforms import validators, ValidationError

class SSGeneratorForm(Form):
    all_colleagues = BooleanField("Generate for all colleagues")
    emp_id = TextField("Employee ID(s)")
    emp_db_id = TextField("Employee DB Spreadsheet ID",[validators.Required("Please enter spreadsheet ID")])
    ss_id = TextField("Salary Spreadsheet ID",[validators.Required("Please enter spreadsheet ID")])
    ws_name = TextField("Worksheet Name",[validators.Required("Please enter worksheet name")])
    gen_pdf = BooleanField("Generate PDF")
    auto_email = BooleanField("Auto send emails")
    email = TextField("Sender Email",[validators.Required("Please enter your email")])
    pwd = PasswordField("Email Password",[validators.Required("Please enter your email password")])
    submit = SubmitField("Generate")