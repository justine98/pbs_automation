from BaseFileOperator import BaseFileOperator 
from datetime import datetime
import os
import xlrd

from transform.MailDrop48HoursTransform import MailDrop48HoursTransform
from transform.MailDropOverpaymentTransform import MailDropOverpaymentTransform

def validate_date_format(date_string, date_format='%Y_%m_%d'):
    try:
        is_valid_date = bool(datetime.strptime(date_string, date_format))
    except ValueError:
        is_valid_date = False

    return is_valid_date

class PBSFileOperator(BaseFileOperator):
    base_folder = "MNL Paygroup/Oncycle"
    
    input_file_pattern_list = [
        '*.xls'
    ]

    input_file_operators = {
        'PQ0032CS__48_HOURS': MailDrop48HoursTransform,
        'PY_OVERPAYMENT_WAGE_ANALYSIS': MailDropOverpaymentTransform,
        'AM_PLA_PAYLINE_E': None,
        'MLA_AM_DETAIL_ADV_VAC': None,
        'HOURLY_EXEMPT_CHANGE_DLOWE': None
    }

    def process_file(self):
        base_folder = self.base_folder
        input_file_path = self.input_file_path
        input_file_operators = self.input_file_operators

        xls = xlrd.open_workbook(input_file_path, on_demand=True)
        sheet_names = xls.sheet_names()
        if "Original" in sheet_names or "Working" in sheet_names:
            return
        else:
            print("New File to Process: %s" % input_file_path) 

        input_file_key_list = list(input_file_operators.keys())

        sub_folder_list = os.listdir(self.base_folder)
        date_folder_list = [sub_folder for sub_folder in sub_folder_list if validate_date_format(sub_folder)]

        if not date_folder_list:
            raise Exception(f"No Valid Date Folder Found in {base_folder}")
        latest_date_folder_name = max(date_folder_list)
        print(latest_date_folder_name)

        try:
            if latest_date_folder_name not in input_file_path:
                print("File is not in latest date folder")
                raise
            else:
                print("File is in latest date folder")

            matched_file_id_index = list(map(input_file_path.__contains__, input_file_key_list)).index(True)
        except:
            print("New File is not for processing.")
            return

        print("Matching File Identifier: ")
        input_file_key = input_file_key_list[matched_file_id_index]
        print(input_file_key)

        pbs_transform_class = input_file_operators[input_file_key]

        if pbs_transform_class is not None:
            print(f"Extending {pbs_transform_class.__name__}")
            pbs_transform = pbs_transform_class(input_file_path)
            pbs_transform.start_transformation()