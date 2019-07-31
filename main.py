import gspread
import numpy as np

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

    def get_ss(self, sheet_id):
        credentials = ServiceAccountCredentials.\
                from_json_keyfile_name(LOCAL_FOLDER + GSPREAD_JSON_FILENAME, SCOPES)
        gc = gspread.authorize(credentials)
        self.ss = gc.open_by_key(sheet_id)

    def get_ws_rng(self, ws, strt_cell):
        rng = None
        height = None
        width = None
        if self.ss is not None:
            match = re.match(r"([a-z]+)([0-9]+)", strt_cell, re.I)
            header_row_len = len(ws.row_values(match.groups()[1]))
            col_number = LETTERS.index(match.groups()[0].lower())
            first_col_len = len(ws.col_values(col_number+1))
            width = header_row_len
            height = first_col_len - int(match.group()[1]) + 1
            if header_row_len < len(LETTERS):
                rng_str = strt_cell + ':' + \
                    LETTERS[header_row_len-1].upper() + str(first_col_len)
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
        loc = np.where(headers == col_name)[1][0]
        col = arr[:, loc]
        return col

    # where 'row_name' below is the item in the first column
    # of the desired row
    @staticmethod
    def get_row_from_arr(arr, row_name):
        loc = np.where(arr[:, 0] == row_name)[0][0]
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