from __future__ import print_function
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
import os
import re


LOCAL_FOLDER = os.path.dirname(os.path.abspath(__file__))
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://spreadsheets.google.com/feeds']
LETTERS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
           'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
           'aa', 'ab', 'ac', 'ad', 'ae', 'af', 'ag', 'ah', 'ai', 'aj', 'ak',
           'al', 'am', 'an', 'ao', 'ap', 'aq', 'ar', 'as', 'at', 'au', 'av',
           'aw', 'ax', 'ay', 'az']
GSPREAD_JSON_FILENAME = "/gspread.json"


# This class wraps functionality to create a spreadsheet object,
# rerieve the first worksheet in the spreadsheet, retrieve a
# range from the worksheet, and convert the range into an 2d
# numpy array for further processing.
class SheetManager:

    def get_first_ws(self):
        credentials = ServiceAccountCredentials.\
                from_json_keyfile_name(LOCAL_FOLDER + GSPREAD_JSON_FILENAME, SCOPES)
        gc = gspread.authorize(credentials)
        self.first_ws = self.ss.sheet1
        return self.first_ws

    def get_ws_by_name(self, ws_name):
        credentials = ServiceAccountCredentials.\
                from_json_keyfile_name(LOCAL_FOLDER + GSPREAD_JSON_FILENAME, SCOPES)
        gc = gspread.authorize(credentials)
        ws = self.ss.worksheet(ws_name)
        return ws

    def get_ss(self, sheet_id):
        print(os.path.abspath(os.path.dirname(__file__)))
        credentials = ServiceAccountCredentials.\
                from_json_keyfile_name(LOCAL_FOLDER + GSPREAD_JSON_FILENAME, SCOPES)
        gc = gspread.authorize(credentials)
        self.ss = gc.open_by_key(sheet_id)

    def get_ws_rng(self, ws, strt_cell, height=0):
        rng = None
        width = None
        if self.ss is not None:
            match = re.match(r"([a-z]+)([0-9]+)", strt_cell, re.I)
            header_row_len = len(ws.row_values(match.groups()[1]))
            col_number = LETTERS.index(match.groups()[0].lower())
            max_col_len = 0
            if height == 0:
                col_lens_arr = []
                for i in range(0, header_row_len):
                    col_len = len(ws.col_values(col_number+i+1))
                    col_lens_arr.append(col_len)
                max_col_len = max(col_lens_arr)
            else:
                max_col_len = height
            width = header_row_len
            height = max_col_len - int(match.groups()[1]) + 1
            if header_row_len < len(LETTERS):
                rng_str = strt_cell + ':' + \
                    LETTERS[header_row_len-1].upper() + str(max_col_len)
                rng = ws.range(rng_str)
            else:
                print('Unable to construct range.')
        else:
            print('Spreadsheet object is null.')
        return rng, width, height

    def convert_rng_2d_numpy(self, rng, width, height):
        arr = np.array([])
        for j in range(0, height):
            head_chop = rng[:width]
            head_chop = [x.value for x in head_chop]
            if arr.size == 0:
                arr = np.array([head_chop])
            else:
                arr = np.append(arr, [head_chop], axis=0)
            if len(rng) > width:
                rng = rng[width:]
        return arr

    def __init__(self, sheet_id):
        self.ss = None
        self.get_ss(sheet_id)
        self.first_ws = self.get_first_ws()


# A set of utilities to work with data in a 2D numpy array
# This set of tools assumes that the data in the worksheet
# is in a rectangular format with a top header of column
# names and data below it.
#
# header :  the top row of column names which can be on any
#           row in the worksheet
# arr :     the rectangular array of data below the header
#
class ArrayManager2D:

    # retrieve the entire column by providing the top 1d
    # array of header names, the 2d data array, and the
    # name of the column
    @staticmethod
    def get_column_from_arr(headers, arr, col_name):
        loc = np.where(headers == col_name)[0][0]
        col = arr[:, loc]
        return col

    # where 'row_name' below is the item in the first column
    # of the desired row
    @staticmethod
    def get_row_from_arr(arr, row_name):
        loc = np.where(arr[:, 0] == row_name)[0][0]
        row = arr[loc]
        return row

    @staticmethod
    def get_val_in_row_by_col_name(row, headers, col_name):
        val = None
        try:
            col_num = np.where(headers == col_name)[0][0]
            val = row[col_num]
        except:
            pass
        return val

    # where 'row name' is the item in the specified column
    # of the desired row
    @staticmethod
    def get_row_from_arr_by_col(arr, headers, row_name, col_name):
        col_num = np.where(headers == col_name)[0][0]
        loc = np.where(arr[:, col_num] == row_name)[0][0]
        row = arr[loc]
        return row

    # get a single value in the table using the column name,
    # the name of the first item in the row, the 1d header
    # array, and the 2d data array
    @staticmethod
    def get_intersection(headers, arr, col_name, row_name):
        intersection_value = None
        try:
            col_loc = np.where(headers == col_name)[1][0]
            row_loc = np.where(arr[:, 0] == row_name)[0][0]
            intersection_value = arr[row_loc, col_loc]
        except:
            pass
        return intersection_value

    @staticmethod
    def get_intersection_two_columns(headers, arr, col_name1, col_name2, col1_row_name):
        intersection_value = None
        col_loc = np.where(headers == col_name2)[0][0]
        col1_loc = np.where(headers == col_name1)[0][0]
        row_loc = np.where(arr[:, col1_loc] == col1_row_name)[0][0]
        intersection_value = arr[row_loc, col_loc]
        return intersection_value

    # get a consolidated new 2d array from the parent array
    # using the headers and the required column names
    @staticmethod
    def get_arr_from_col_names(headers, data_arr, col_names):
        final_arr = np.array([])
        final_headers = np.array([])
        for col_name in col_names:
            col = ArrayManager2D.get_column_from_arr(headers, data_arr, col_name)
            if final_arr.size == 0:
                final_arr = np.array([col])
            else:
                final_arr = np.append(final_arr, [col], axis=0)
        final_arr = np.transpose(final_arr)
        return final_arr

    def __init__(self, ws):
        pass


def get_ss(sheet_key):
    credentials = ServiceAccountCredentials.\
            from_json_keyfile_name(LOCAL_FOLDER + GSPREAD_JSON_FILENAME, SCOPES)
    gc = gspread.authorize(credentials)
    ws = gc.open_by_key(sheet_key)
    return ws


# get the 2 dimensional container of all colleagues data from
# a worksheet in the main salary spreadsheet
def get_2d_container_from_ws(ws, start_cell_a1_name, start_cell_row, colleague_col):
    container = []
    top_header_row_len = len(ws.row_values(start_cell_row))
    num_colleagues = len(ws.col_values(colleague_col))
    if num_colleagues is not 0:
        range_2d_str = start_cell_a1_name + str(start_cell_row) + ":" + \
            str(LETTERS[top_header_row_len-1]).upper() + str(num_colleagues)
        rng = ws.range(range_2d_str)
        tmp_arr = []
        counter = 0
        for val in rng:
            tmp_arr.append(val.value)
            if counter == top_header_row_len - 1:
                container.append(tmp_arr)
                tmp_arr = []
                counter = 0
            else:
                counter += 1
    return container


# get a consolidated list of all colleagues in all the sheets
# in the main salary spreadsheet
def get_consolidated_colleague_list(containers):
    colleagues_list = []
    for container in containers:
        emp_list = []
        counter = 0
        for row in container:
            if counter != 0:
                emp_list.append(row[2])
            counter += 1
        colleagues_list.extend(emp_list)
    colleagues_list = list(dict.fromkeys(colleagues_list))
    return colleagues_list


def split_comma_seperated_str(comma_str):
    split_arr = comma_str.split(',')
    split_arr = [x.strip() for x in split_arr]
    return split_arr


# get the data for a single colleague from the 2 dimensional list
def get_colleague_data_from_container(name, container):
    colleague_data_row = []
    index = 0
    for row in container:
        if row[2] == name:
            break
        index += 1
    colleague_data_row = container[index]
    return colleague_data_row


# get all salary slip worksheets in the main spreadsheet
def get_salary_slip_sheets(main_sheet):
    ws = main_sheet.worksheets()
    return ws


# prepare the summed up data for the colleague for the required
# fields
def prepare_colleague_summed_data(colleague, containers, required_fields):
    data_arr = []
    for i in range(len(required_fields)):
        data_arr.append(0)
    for container in containers:
        colleague_found = False
        indices_list = []
        for field in required_fields:
            if field in container[0]:
                indices_list.append(container[0].index(field))
            else:
                indices_list.append(0)
        colleague_index = 0
        for row in container:
            if row[2] == colleague:
                colleague_found = True
                break
            colleague_index += 1
        for i in range(len(data_arr)):
            # remove comma if any
            index_val = indices_list[i]
            if index_val is not 0:
                if colleague_found == True:
                    val = str(container[colleague_index][indices_list[i]]).replace(',', '')
                else:
                    val = 0
                if val == '':
                    val = 0
            else:
                val = 0
            data_arr[i] += int(val)
    return data_arr


# print the data passed into a new worksheet using the range
def print_data_to_ws(ss, rng_str, data_container):
    ws = None
    try:
        ws = ss.worksheet('summed data')
    except:
        ws = ss.add_worksheet('summed data', 1000, 50)
    cell_list = ws.range(rng_str)
    counter = 0
    for cell in cell_list:
        cell.value = data_container[counter]
        counter += 1
    ws.update_cells(cell_list)


def convert_number_to_words(number):
    converted = ''
    try:
        digits_map = {'1': 'One ', '2': 'Two ', '3': 'Three ', '4': 'Four ',
                      '5': 'Five ', '6': 'Six ', '7': 'Seven ', '8': 'Eight ',
                      '9': 'Nine ', '0': ''}
        tens_map = {'2': 'Twenty ', '3': 'Thirty ', '4': 'Fourty ',
                    '5': 'Fifty ', '6': 'Sixty ', '7': 'Seventy ', '8': 'Eighty ',
                    '9': 'Ninety ', '0': ''}
        teens = {'11': 'Eleven ', '12': 'Twelve ', '13': 'Thirteen ',
                 '14': 'Fourteen ', '15': 'Fifteen ', '16': 'Sixteen ',
                 '17': 'Seventeen ', '18': 'Eighteen ', '19': 'Nineteen '}
        num_to_str = str(number)
        num_to_str = ''.join(num_to_str.split(','))
        str_len = len(num_to_str)
        str_counter = 0
        for i in range(str_len, 0, -1):
            if i == 7:
                converted += digits_map[num_to_str[str_counter]]
                converted += 'Million '
            if i == 6 or i == 3:
                val = num_to_str[str_counter]
                if val is not '0':
                    converted += digits_map[num_to_str[str_counter]]
                    converted += 'Hundred '
            if i == 5 or i == 2:
                val = num_to_str[str_counter]
                if val is not '1':
                    converted += tens_map[num_to_str[str_counter]]
                else:
                    next_val = num_to_str[str_counter + 1]
                    concat = num_to_str[str_counter] + next_val
                    converted += teens[concat]
            if i == 4:
                val = num_to_str[str_counter]
                prev_val = num_to_str[str_counter - 1]
                if val is not '0':
                    if prev_val is not '1':
                        converted += digits_map[num_to_str[str_counter]]
                if str_len == 7:
                    prev_val_2 = num_to_str[str_counter - 2]
                    if val == '0' and prev_val == '0' and prev_val_2 == '0':
                        pass
                    else:
                        converted += 'Thousand '
                else:
                    converted += 'Thousand '
            if i == 1:
                prev_val = num_to_str[str_counter - 1]
                if prev_val is not '1':
                    converted += digits_map[num_to_str[str_counter]]
                converted += 'Dollars Only'

            str_counter += 1

    except:
        print("Unable to generate in words for one case.")

    return converted
