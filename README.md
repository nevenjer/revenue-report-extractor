import os
from openpyxl import load_workbook, Workbook

def create_output_file(output_file):
    # Create a new file and rename the sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "data_cs_bpm"
    
    # Set the header in row 1.
    headers = ['id', 'num_bill', 'num_clients', 'id_pos', 'ca', 'name', 'service', 
               'recipt', 'cash', 'cheque', 'pay_in_slip', 'qr_payment', 'date_time', 'rounding_num']
    ws.append(headers)
    
    # Save file
    wb.save(output_file)

def process_file(input_file, output_file):
    wb_input = load_workbook(input_file)
    ws_input = wb_input.active

    if ws_input['K17'].value == 'รวม':
        return

    if os.path.exists(output_file):
        wb_output = load_workbook(output_file)
    else:
        wb_output = Workbook()
        ws_output = wb_output.active
        ws_output.title = "data_cs_bpm"

    ws_output = wb_output['data_cs_bpm']

    # Start writing data from row 2
    start_row = ws_output.max_row + 1

    for row in range(13, ws_input.max_row + 1):  # Starting from row 13 to the last row.
        # Check if K17 has the word 'รวม' in this row.
        if ws_input[f'K{row}'].value == 'รวม':
            continue

        # Pull data from each column
        data = {
            'id': start_row - 1,
            'num_bill': ws_input[f'D{row}'].value,
            'num_clients': ws_input[f'E{row}'].value,
            'id_pos': ws_input[f'F{row}'].value,
            'ca': ws_input[f'J{row}'].value,
            'name': ws_input[f'K{row}'].value,
            'service': ws_input[f'O{row}'].value,
            'recipt': ws_input[f'P{row}'].value,
            'cash': ws_input[f'R{row}'].value,
            'cheque': ws_input[f'T{row}'].value,
            'pay_in_slip': ws_input[f'V{row}'].value,
            'qr_payment': ws_input[f'M{row}'].value,
            'date_time': ws_input[f'N{row}'].value,
            'rounding_num': ws_input[f'C{row}'].value
        }

        # Write data to output file
        ws_output.append(list(data.values()))  # Add data to the next row

        # Add automatic number
        ws_output[f'A{start_row}'] = start_row - 1
        ws_output[f'B{start_row}'] = 1 if ws_input[f'D{row}'].value == 1 else 0
        ws_output[f'C{start_row}'] = 1 if ws_input[f'C{row}'].value == 1 else 0

        start_row += 1  # Add a row for the next record.

    # Save the output file
    wb_output.save(output_file)

# Specify the directory to work in.
output_file = r'D:\Your File Path\'

# Create a new output file
create_output_file(output_file)

# Loop through files in FREQ_10 folder
for i in range(671008, 671011):  # Rename the specified files to the desired number.
    input_file = f'D:\Your File Path\'
    if os.path.exists(input_file):
        process_file(input_file, output_file)
