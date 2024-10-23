import os
from openpyxl import load_workbook, Workbook

def create_output_file(output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "data_cs_bpm"

    headers = ['no.', 'id_pos', 'ca', 'name', 'service', 
               'recipt', 'debt', 'rounding_num', 'total', 
               'cash', 'cheque', 'pay_in_slip', 'qr_payment', 
               'date_time', 'status', 'num_bill', 'num_clients']
    ws.append(headers)
    wb.save(output_file)
    
def process_file(input_file, output_file):
    wb_input = load_workbook(input_file)
    ws_input = wb_input.active

    if os.path.exists(output_file):
        wb_output = load_workbook(output_file)
    else:
        wb_output = Workbook()
        ws_output = wb_output.active
        ws_output.title = "data_cs_bpm"

    ws_output = wb_output['data_cs_bpm']
    
    start_row = ws_output.max_row + 1

    for row in range(13, ws_input.max_row + 1):
        data = { 
            'no': ws_input[f'C{row}'].value,                          
            'id_pos': ws_input[f'D{row}'].value,
            'ca': ws_input[f'E{row}'].value,
            'name': ws_input[f'F{row}'].value,
            'service': ws_input[f'J{row}'].value,
            'recipt': ws_input[f'K{row}'].value,
            'debt': ws_input[f'L{row}'].value,
            'rounding_num': ws_input[f'M{row}'].value,
            'total': ws_input[f'N{row}'].value,
            'cash': ws_input[f'O{row}'].value,
            'cheque': ws_input[f'P{row}'].value,
            'pay_in_slip': ws_input[f'R{row}'].value,
            'qr_payment': ws_input[f'T{row}'].value,
            'date_time': ws_input[f'V{row}'].value,            
            'status': ws_input[f'X{row}'].value
        }
    
        ws_output.append(list(data.values()))

    # Remove rows with 'รวม' in column F
    for row in range(ws_output.max_row, 1, -1):  # Start from the bottom
        if ws_output[f'F{row}'].value and str(ws_output[f'F{row}'].value).startswith('รวม'):
            ws_output.delete_rows(row)

    # Fill 'num_bill' with 1 starting from row 2
    for row in range(2, ws_output.max_row + 1):
        ws_output[f'P{row}'] = 1  # Assuming 'num_bill' is the 16th column

    # Fill 'num_clients' with 1 for rows with data in column A
    for row in range(2, ws_output.max_row + 1):
        if ws_output[f'A{row}'].value:
            ws_output[f'Q{row}'] = 1  # Assuming 'num_clients' is the 17th column

    wb_output.save(output_file)

output_file = r'D:\Your Path File.xlsx'
create_output_file(output_file)

for i in range(671008, 671017):  # Rename the specified files to the desired number.
    input_file = f'D:\Your Path File.xlsx'
    if os.path.exists(input_file):
        process_file(input_file, output_file)

print('----------Data extraction from your report completed successfully----------')
