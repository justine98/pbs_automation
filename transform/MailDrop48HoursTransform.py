from BasePBSTransform import BasePBSTransform

import pandas as pd
import xlwings as xw
import xlsxwriter

class MailDrop48HoursTransform(BasePBSTransform):
    def __init__(self, input_file_path):
        super().__init__(input_file_path)

    def start_transformation(self):
        print("Starting Transformation for Mail Drop: 48 Hours")

        input_file_path = self.input_file_path
        header_index = 3
        input_df = pd.read_excel(input_file_path, sheet_name="Working", header=header_index)

        print(input_df.columns)

        header_data_df = pd.read_excel(input_file_path, sheet_name="Working", usecols='A', skiprows=1, nrows=2, header=None)

        pay_period = header_data_df.iloc[0][0].split("=")[-1].strip()
        pay_group = header_data_df.iloc[1][0].split("=")[-1].strip()
        print(pay_period)
        print(pay_group)

        # input_df['Pay Period'] = pay_period
        # input_df['Pay Group'] = pay_group

        validation_data_df = input_df.loc[:, ["Empl ID", "Earn Code", "Sep Check Nbr"]]
        validation_data_dict = validation_data_df.to_dict('records')
        print(pay_period, pay_group, validation_data_dict)

        validatation_result_df = pd.DataFrame.from_dict(validation_data_dict)
        validatation_result_df['Action Done'] = validatation_result_df["Empl ID"] 

        column_index = len(input_df.columns)
        column_letter = xlsxwriter.utility.xl_col_to_name(column_index)
        # column_location = f'{column_letter}{header_index}'
        # print(column_location)
        header_number = header_index + 1

        work_book = xw.books.open(input_file_path)
        working_sheet = work_book.sheets["Working"]

        # working_sheet[f'A{header_number}'].copy()
        # working_sheet[f'{column_letter}{header_number}'].paste("formats")

        working_sheet[f'{column_letter}{header_number}'].options(index=False).value = validatation_result_df['Action Done']
        # working_sheet[f'{column_letter}{header_number + 1}'].value = validatation_result_df['Action Done']
        print(working_sheet.range(f'{column_letter}{header_number + 1}').expand('down'))
        working_sheet.range(f'{column_letter}{header_number + 1}').expand('down').color = (255,255,0)
        work_book.save()
        work_book.close()

    # def populate_column():

        





