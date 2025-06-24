from openpyxl.utils import coordinate_to_tuple

def validate_cell_value(cell, value_type):
    if cell.value is None:
        return None, True
    
    if value_type == 'v-number':
        if isinstance(cell.value, int):
            return cell.value, True
    
        try:
            return int(cell.value), True
        except:
            pass
    
    if value_type == 'v-string':    
        return cell.value, True
    
    return cell.value, False

def copy_cell_value(master_ws, master_row, master_col, input_ws, input_row, input_col, value_type, logger, input_file):
    input_ws_cell = input_ws.cell(input_row, input_col)
    input_ws_cell_value, is_valid = validate_cell_value(input_ws_cell, value_type)
    if is_valid:
        master_ws.cell(master_row, master_col).value = input_ws_cell_value
        
    else:
        logger.error(f'{input_file} | Wrong value of cell: {input_ws.cell(input_row, input_col).coordinate} | [{input_ws_cell.value}]')

def get_keys_in_worksheet(ws):
    keys = {}

    # get number of rows
    num_rows = ws.max_row

    # loop through each row
    for row in range(1, num_rows + 1):
        if ws.cell(row, 1).value:
            if not ws.cell(row,2).font.bold:
                keys[ws.cell(row, 1).value] = {
                                                'row': row,
                                                'row_type': 'detail'
                                                }
            else:
                keys[ws.cell(row, 1).value] = {
                                                'row': row,
                                                'row_type': 'summary'
                                                }
    try:
        keys.pop('TOTAL DAV')
    except:
        pass

    return keys