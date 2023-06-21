import openpyxl
from pathlib import Path
import os

path = Path('./data')

ir = r'data\\CYP_DE Asset Tagging requirement.xlsx'
amr = r'data\\TAS-CYP-CBS-ZWD-REG-XIM-NAP-X0001_C.xlsx'

amr_data = openpyxl.load_workbook(amr)

ir_data = openpyxl.load_workbook(ir)

amr_valiodation_ws = amr_data.get_sheet_by_name('Validation Tables')
ir_mep_data = ir_data.get_sheet_by_name('MEP Data Requirement')

cyp_types = ir_mep_data.tables['CYPTYPES']

#print([os.fspath(x) for x in path.iterdir() if x.suffix == '.xlsx'])
#print(cyp_types.column_names[10])
#print(cyp_types.tableColumns['Type Code Status'])

#print(len(cyp_types.tableColumns-1))

for i in range(0,(len(cyp_types.tableColumns))):
    #print(cyp_types.column_names[i])
    #print(cyp_types.tableColumns[i].name)
    if cyp_types.column_names[i] == 'Scheduled in AMR':
        print('Found it')
        print(i)
    else:
        continue
