import openpyxl as xl
import os
from init.app_logger import *
from init.lib import *

#init
sharepoint_folder = r'D:\temp\OneDrive_2025-07-02\Forecast meeting\1. FY25'
sharepoint_folder = sharepoint_folder.replace('\\', '/')
logger.debug(f'Sharepoint folder: "{sharepoint_folder}"')

# open the workbook
master_file = get_master_file(sharepoint_folder, None)
logger.debug(f'Master file: "{master_file}"')

if not master_file:
    logger.error('Master file not found')
    exit()

master_wb = xl.load_workbook(master_file, data_only=True)
ws = master_wb.active

# get all keys from template file
keys_template = get_keys_in_worksheet(ws) # { key: row }

# get columns to copy
columns_to_copy = [('JK', 'v-number'),    # Result JUN_25
                ('JW', 'v-string'),    # Result Review
                ('KC', 'v-number'),    # Topline 7/3
                ('KL', 'v-string'),    # topline review
                ('KM', 'v-string')]    # Action-plan

# get all file .xlsx in input folder
input_files = get_input_files(sharepoint_folder, None, logger)
logger.debug(f'Total {len(input_files)} files in input folder')
for file in input_files:
    logger.debug(f'>>>>>> "{file}"')


# loop through each file in input folder
for input_file in input_files:
    logger.debug(f'Processing: "{input_file}"')
    # open the input file
    input_wb = xl.load_workbook(input_file, data_only=True)
    input_ws = input_wb.active

    # get keys of input file
    keys_input = get_keys_in_worksheet(input_ws)

    # loop from start_row to total_rows
    for key_id in keys_input:
        for column in columns_to_copy:
            key_object = keys_input[key_id]

            if key_object['row_type'] == 'summary' and column[1] == 'v-number':
                continue
            else:
                column_index =  coordinate_to_tuple(column[0] + '1')[1]

                master_row = keys_template[key_id]['row']
                input_row = keys_input[key_id]['row']

                copy_cell_value(ws, master_row, column_index, input_ws, input_row, column_index, column[1], logger, input_file)

export_file_name = f'output/{timestamp} Topline FCT (combined).xlsx'
master_wb.save(export_file_name)
logger.debug(f'Exported: {export_file_name}')
logger.debug('Script Done!')
