import os
from openpyxl import load_workbook, Workbook

def create_output_file(output_file):
    # Create a new file and change the name of the worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "data_cs_bpm"

    # Set the header in row 1
    headers = ['no.', 'id_pos', 'ca', 'name', 'service', 
               'recipt', 'debt', 'rounding_num', 'total', 
               'cash', 'cheque', 'pay_in_slip', 'qr_payment', 
               'date_time', 'status', 'num_bill', 'num_clients']
    ws.append(headers)
    
    # Save the file
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
    
    # Start writing information from row 1
    start_row = ws_output.max_row + 1

    for row in range(13, ws_input.max_row + 1):  # Starting from the 13th row to the last row
        

        # Retrieve data from each column
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
    
        # Write data to output file
        ws_output.append(list(data.values()))  # Add data to the next row

    # Save the output file
    wb_output.save(output_file)

# Specify the directory to work in.
output_file = r'D:\Your Path File\'

# Create a new output file
create_output_file(output_file)

# Loop through files in your folder
for i in range(671008, 671017):  # Rename the specified files to the desired number.
    input_file = f'D:\Your Path File\'
    if os.path.exists(input_file):
        process_file(input_file, output_file)

print('The data extraction operation in your report was successful.')
