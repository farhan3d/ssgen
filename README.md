# ssgen
A minimal experimental Google Sheets, Flask, and Heroku based salary slip generation and delivery system.

The following need to be pip installed and added to the requirements.txt for Heroku:

gunicorn
flask
flask-wtf
gspread
pandas
premailer
numpy
oauth2client
flask-login

Working with gspread requires a service token through a new project created on the Google Developers Console. The gspread setup can be followed in the link below:

https://gspread.readthedocs.io/en/latest/

Two spreadsheet databases need to be set-up for the salary slip generation system to work. The format is shown in the below example sheets:

Salary Slip Database: https://docs.google.com/spreadsheets/d/120NOykXIUyKvrM5SSPKK6CzPSOYSk1sv0eREajwfZyU/edit#gid=1285578064
Employee Database: https://docs.google.com/spreadsheets/d/1YYaaDksBX7zwtH17mPudKnoX4yg-K4-ROiigadGQe0s/edit#gid=0

The system delivers salary slips to the employees listed in the database via HTML formatted email by adding the required information in the Flask form.
