import openpyxl as xl
import os
from init.app_logger import *
from init.lib import *

#init
master_file = 'template/master_fct.xlsx'
input_folder = 'input'

# open the workbook
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
input_files = os.listdir(input_folder)
input_files = [file for file in input_files if '~$' not in file]
input_files.remove('.gitkeep')

logger.debug(f'Total {len(input_files)} files in input folder')

# loop through each file in input folder
for input_file in input_files:
    logger.debug(f'Processing: {input_file}')
    # open the input file
    input_file_path = os.path.join(input_folder, input_file)
    input_wb = xl.load_workbook(input_file_path, data_only=True)
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