from __future__ import print_function
from jinja2 import Environment, FileSystemLoader
import os, sys
import os.path
import gspread
import smtplib
import pandas as pd
import time
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import utils
from premailer import transform


SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://spreadsheets.google.com/feeds']
SALARY_SHEET_MONTH_HEADER_CELL = [1, 1]
SALARY_SHEET_CORNER = 'A2'
PEOPLE_SHEET_CORNER = 'A1'

# Any new column additions need to be added in the lists below in exactly
# the same case.

# actual display on salary sheet | name of the column in the sheet
EARNINGS_MAP = {
    'Gross Salary':                 'Actual Gross Pay',
    'Project Bonus':                'Project Bonus',
    'Referral Bonus':               'Referral Bonus',
    'Arrears':                      'Arrears',
    'Leave Encashment':             'Leave Encashment',
    'Conveyance Allowance':         'Conveyance Allow',
    'Special Allowance':            'Special Allow',
    'Total Gross Salary':           'Earned Gross Pay'
}
# actual display on salary sheet | name of the column in the sheet
DEDUCTIONS_MAP = {
    'Income Tax':                   'Income Tax Deduction at source',
    'Food':                         'Advance against Salary(Food)',
    'Fines':                        'Advance against Salary(Fines)',
    'Advance / Other Deduction':    'Other Deductions',
    'Total Deductions':             'Total',
    'Net Salary':                   'Net Pay'
}
# actual display on salary sheet | name of the column in the sheet
BIO_MAP = {
    'Designation':                  'Designation',
    'Total Working Days':           'Total Working Days',
    'Days Worked':                  'Actual Working Days',
    'Employee Name':                'Name Of Employee'
}

LOCAL_FOLDER = os.path.dirname(os.path.abspath(__file__))
# setting below as 0 would slow down the process because the program would be
# reading every column to get the one with the most rows
EMPLOYEE_COUNT_OVERRIDE = 1000

app = None

path = os.path.dirname(os.path.abspath(__file__))
env = Environment(loader=FileSystemLoader(path + '/templates'))
template = env.get_template("myreport.html")
COLLEAGUE_COMPANY_INFO = ['ACME Industries', 'ACME Street 911', 'USA']


def get_ws(sheet_key):
    credentials = ServiceAccountCredentials.\
            from_json_keyfile_name(LOCAL_FOLDER + "/gspread.json", SCOPES)
    gc = gspread.authorize(credentials)
    ws = gc.open_by_key(sheet_key).sheet1
    return ws


def get_ss(sheet_key):
    credentials = ServiceAccountCredentials.\
            from_json_keyfile_name(LOCAL_FOLDER + "/gspread.json", SCOPES)
    gc = gspread.authorize(credentials)
    ws = gc.open_by_key(sheet_key)
    return ws


def toggle_email_inputs(dummy=None):
    if app.get_checkbox() == 1:
        app.get_email_input_obj().entry.config(state=NORMAL)
        app.get_email_pwd_input_obj().entry.config(state=NORMAL)
    elif app.get_checkbox() == 0:
        app.get_email_input_obj().entry.config(state=DISABLED)
        app.get_email_pwd_input_obj().entry.config(state=DISABLED)


def toggle_emp_id_input(dummmy=None):
    if app.get_all_emp_checkbox() == 0:
        app.get_emp_id_input().entry.config(state=NORMAL)
    elif app.get_all_emp_checkbox() == 1:
        app.get_emp_id_input().entry.config(state=DISABLED)


def get_current_month_header(sheet_id, ws_name):
    # salary_sheet = get_ws(sheet_id)
    ss = get_ss(sheet_id)
    salary_sheet = ss.worksheet(ws_name)
    cell = salary_sheet.cell(SALARY_SHEET_MONTH_HEADER_CELL[0],
                             SALARY_SHEET_MONTH_HEADER_CELL[1])
    return cell.value


def get_inputs():
    emp_id_str = app.get_emp_id()
    emp_ids = utils.split_comma_seperated_str(emp_id_str)
    sheet_id = app.get_salary_sheet_id()
    ws_id = app.get_ws_id()
    checkbox = app.get_checkbox()
    emp_checkbox = app.get_all_emp_checkbox()
    if sheet_id and ws_id:
        return [emp_ids, sheet_id, ws_id, checkbox, emp_checkbox]
    else:
        return None


def get_email_inputs():
    email_input = app.get_email()
    pwd_input = app.get_email_pwd()
    if email_input and pwd_input:
        return [email_input, pwd_input]
    else:
        return None


# generate pandas dataframe for html output
def create_salary_dataframe(headers, data):
    dataframe_dict = {}
    header = ['Earnings', 'Amount Earned', 'Deductions', 'Amount Deducted']
    earnings_dict_iter = iter(EARNINGS_MAP)
    col1_data = []
    col2_data = []
    for row in earnings_dict_iter:
        col1_data.append(row)
        col2_data.append(
            utils.ArrayManager2D.get_val_in_row_by_col_name(
                data, headers, EARNINGS_MAP[row]
            )
        )
    earnings_dict_iter = iter(DEDUCTIONS_MAP)
    col3_data = []
    col4_data = []
    for row in earnings_dict_iter:
        col3_data.append(row)
        col4_data.append(
            utils.ArrayManager2D.get_val_in_row_by_col_name(
                data, headers, DEDUCTIONS_MAP[row]
            )
        )
    if len(col1_data) < len(col3_data):
        for i in range(0, len(col3_data) - len(col1_data)):
            col1_data.append('')
            col2_data.append('')
    elif len(col1_data) > len(col3_data):
        for i in range(0, len(col1_data) - len(col3_data)):
            col3_data.append('')
            col4_data.append('')
    dataframe_dict = {header[0]: col1_data, header[1]: col2_data,
                      header[2]: col3_data, header[3]: col4_data}
    df = pd.DataFrame(dataframe_dict, columns=[header[0], header[1],
                                               header[2], header[3]])
    return df


def create_bio_dataframe(headers, data, jd=''):
    col1_data = ['Designation', 'Date of Joining', 'Employment Status',
                 'Total Working Days', 'Days Worked']
    designation = utils.ArrayManager2D.get_val_in_row_by_col_name(
        data, headers, BIO_MAP[col1_data[0]]
    )
    working_days = utils.ArrayManager2D.get_val_in_row_by_col_name(
        data, headers, BIO_MAP[col1_data[3]]
    )
    days_worked = utils.ArrayManager2D.get_val_in_row_by_col_name(
        data, headers, BIO_MAP[col1_data[4]]
    )
    col2_data = [designation, jd, 'Full-time', working_days, days_worked]
    data_tuples = list(zip(col1_data, col2_data))
    emp_name = utils.ArrayManager2D.get_val_in_row_by_col_name(
        data, headers, BIO_MAP['Employee Name']
    )
    df = pd.DataFrame(data_tuples, columns=['Employee Name', emp_name])
    return df


def create_salary_in_words_dataframe(salary_in_words):
    df_dict = {}
    df = pd.DataFrame(df_dict, columns=['Net salary in words',
                                        salary_in_words])
    return df


def salary_slip_as_html(colleague_name, colleague_email,
                        email_html, email_inputs,
                        month_information=None):
    msg = MIMEMultipart()
    msg['From'] = email_inputs[0]
    msg['To'] = ", ".join([colleague_email])
    subject = ''
    if month_information == None:
        subject = 'Salary Slip for ' + colleague_name
    else:
        subject = month_information
    msg['Subject'] = subject
    msg.attach(MIMEText(email_html, 'html'))
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_inputs[0], email_inputs[1])
    text = msg.as_string()
    server.sendmail(email_inputs[1], [colleague_email], text)
    server.quit()


def get_data_from_people_sheet(sheet_id):
    people_sheet_dict = {}
    ps = utils.SheetManager(sheet_id)
    ps_rng, width, height = ps.get_ws_rng(ps.first_ws, PEOPLE_SHEET_CORNER,
        EMPLOYEE_COUNT_OVERRIDE)
    arr = ps.convert_rng_2d_numpy(ps_rng, width, height)
    for row in arr:
        emp_name = utils.ArrayManager2D.get_val_in_row_by_col_name(
            row, arr[0], 'Employee Name'
        )
        emp_email = utils.ArrayManager2D.get_val_in_row_by_col_name(
            row, arr[0], 'Email'
            )
        emp_id = utils.ArrayManager2D.get_val_in_row_by_col_name(
            row, arr[0], 'Code'
            )
        people_sheet_dict[emp_name] = {'email': emp_email, \
                                       'id': emp_id}
    return people_sheet_dict


def generate_sent_confirmation_email(email_inputs):
    msg = MIMEMultipart()
    msg['From'] = email_inputs[0]
    msg['To'] = ", ".join([email_inputs[0]])
    subject = 'Salary slips processing completion'
    msg['Subject'] = '[automsg] ' + subject

    body = "Processing for salary slips is complete. \
            Please refer to application log for any issues."
    msg.attach(MIMEText(body, 'html'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_inputs[0], email_inputs[1])
    text = msg.as_string()
    server.sendmail(email_inputs[1], [email_inputs[0]], text)
    server.quit()


def generate(web_data=None):
    items_arr = [[web_data[0]], web_data[1], web_data[2],
                web_data[3], web_data[4], web_data[5], web_data[6], web_data[7]]
    emp_ids, sheet_id, people_sheet_id, ws_id, checkbox, emp_checkbox = \
            items_arr[0], items_arr[1], items_arr[2], \
            items_arr[3], items_arr[4], items_arr[5]
    emp_ids = utils.split_comma_seperated_str(emp_ids[0])
    sheet_obj = utils.SheetManager(sheet_id)
    ws = sheet_obj.get_ws_by_name(ws_id)
    rng, width, height = sheet_obj.get_ws_rng(
        ws, SALARY_SHEET_CORNER, EMPLOYEE_COUNT_OVERRIDE)
    data_arr = sheet_obj.convert_rng_2d_numpy(rng, width, height)
    css_file = str(open(LOCAL_FOLDER + '/typography.css', 'r').read())
    arr = sheet_obj.convert_rng_2d_numpy(rng, width, height)
    if emp_checkbox:
        # get the entire list of employee id's and merge with id's
        # entered in the input box
        emp_list_full = utils.ArrayManager2D.get_column_from_arr(
            arr[0], arr[1:], 'Code'
        )
        emp_ids = emp_ids + list(set(emp_list_full) - set(emp_ids))
    people_sheet_dict = get_data_from_people_sheet(people_sheet_id)
    for emp_id in emp_ids:
        if emp_id.isdigit():
            emp_row = utils.ArrayManager2D.get_row_from_arr_by_col(
                data_arr, data_arr[0], emp_id, 'Code')
            if emp_row is not None:
                emp_name = utils.ArrayManager2D.get_val_in_row_by_col_name(
                    emp_row, data_arr[0], BIO_MAP['Employee Name']
                )
                if emp_name in people_sheet_dict:
                    emp_jd = ''
                    bio_df = create_bio_dataframe(data_arr[0], emp_row, emp_jd)
                    sal_df = create_salary_dataframe(data_arr[0], emp_row)
                    emp_net_salary = utils.ArrayManager2D.get_val_in_row_by_col_name(
                        emp_row, data_arr[0], DEDUCTIONS_MAP['Net Salary'])
                    net_salary_in_words = utils.convert_number_to_words(emp_net_salary)
                    sal_in_words_df = create_salary_in_words_dataframe(
                        net_salary_in_words)
                    month_information = get_current_month_header(sheet_id, ws_id)
                    template_vars = {"title": COLLEAGUE_COMPANY_INFO[0],
                                    "company_name": COLLEAGUE_COMPANY_INFO[0],
                                    "address_line_1": COLLEAGUE_COMPANY_INFO[1] + ' ' + \
                                        COLLEAGUE_COMPANY_INFO[2],
                                    "month_information": month_information,
                                    "bio_data": bio_df.to_html(index=False),
                                    "salary_data": sal_df.to_html(index=False),
                                    "net_salary_in_words": sal_in_words_df.to_html(
                                        index=False)}
                    html_out = template.render(template_vars)
                    css_file = open(LOCAL_FOLDER + '/typography.css', 'r').read()
                    split_html = html_out.split('</head>')
                    html_out = split_html[0] + '\n' + '</head>' + '\n' + '<style>' + \
                        '\n' + css_file + '\n' + '</style>' + '\n' + split_html[1]
                    html_out = transform(html_out)
                    email_inputs = [items_arr[6], items_arr[7]]
                    if people_sheet_dict:
                        # send the current colleague their salary slip
                        current_colleague_email = people_sheet_dict[emp_name]['email']
                        if checkbox:
                            if email_inputs:

                                if str(emp_id) == \
                                str(people_sheet_dict[emp_name]['id']):
                                    salary_slip_as_html(emp_name, current_colleague_email,
                                        html_out, email_inputs, month_information)
                                    time.sleep(5)
                                    message = "<br>Salary slip sent to: " + emp_name
                                    yield message
                                else:
                                    message = '<br>Unable to send email to ' + \
                                        emp_name + ', ID mismatch.'
                                    yield message
                else:
                    message = "<br>" + emp_name + ' is missing in the employee sheet.'
                    yield message
    yield "<br><br>COMPLETED!"
