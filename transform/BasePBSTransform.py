import xlwings as xw
import pandas as pd
import threading

class BasePBSTransform:
    def __init__(self, input_file_path):
        self.input_file_path = input_file_path
        self.is_valid = True
        self.input_df = None
        self.initialize_sheets()
    
    def initialize_sheets(self):
        input_file_path = self.input_file_path
        print("Initializing Sheets in File")
        app = xw.App(visible=False)
        work_book = xw.books.open(input_file_path)
        sheet_names = work_book.sheet_names

        base_sheet_name = "sheet1"
        copy_sheet_name = "Working"
        new_sheet_name = "Original"

        if base_sheet_name in sheet_names and copy_sheet_name not in sheet_names:
            print(f"Sheets to Prepare: {sheet_names}")
            print(f"Duplicating {base_sheet_name} as {copy_sheet_name}")
            base_sheet = work_book.sheets[base_sheet_name]
            base_sheet.copy(name=copy_sheet_name)
            print(f"Renaming {base_sheet_name} to {new_sheet_name}")
            base_sheet.name = new_sheet_name
            work_book.save()

            sheet_names = work_book.sheet_names
            print(f"Sheets to Validate: {sheet_names}")
            work_book.close()
        else:
            raise Exception("Don't Worry")

    def start_transformation(self):
        if self.is_valid:
            print("Starting Transformation")
        else:
            print("Skip File")