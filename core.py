#from update_active_type_code import update_ir_bool_matrix
from pathlib import Path
from openpyxl import load_workbook
import pandas as pd
from itertools import islice
import re





twin_codes = r"C:\Users\SamuelOnuoha\John Holland Group\CYP-Digital Engineering - Working Folder\dev\data\twin_codes.xlsx"
ir_file = r"C:\Users\SamuelOnuoha\John Holland Group\CYP-Digital Engineering - Working Folder\dev\data\CYP_DE Asset Tagging requirement.xlsx"
output_path = r"C:\Users\SamuelOnuoha\John Holland Group\CYP-Digital Engineering - Working Folder\dev\data\outputs"

twin_sheetname = 'Export'

""" mep_ir_sheetname = 'MEP Data Requirement'
rshp_ir_sheetname = 'RSHP ARC Data Requirement'
aud_ir_sheetname = 'HWW AUD Data Requirement'
hww_ir_sheetname = 'HWW ARC Data Requirement'

 """

ir_sheets = ['MEP Data Requirement', 'RSHP ARC Data Requirement', 'HWW AUD Data Requirement', 'HWW ARC Data Requirement']

#Generating combined twin codes"

twin_code_df = pd.read_excel(twin_codes,'Export')
active_codes = twin_code_df.loc[twin_code_df['Type Description'].notnull()] 
active_codes_list = active_codes.loc[:, 'MM Type'].to_list()

print(len(active_codes_list))

'''
arc_twin_code_df = pd.read_excel(arc_twin_codes,'Export')
active_codes = arc_twin_code_df.loc[arc_twin_code_df['Type Description'].notnull()] 
arc_active_codes_list = active_codes.loc[:, 'MM Type'].to_list()
'''


"""
 # Generating AUD Twin Code list
aud_twin_code_df = pd.read_excel(aud_twin_codes,'Export')
aud_active_codes = aud_twin_code_df.loc[aud_twin_code_df['Type Description'].notnull()] 
aud_active_codes_list = aud_active_codes.loc[:, 'MM Type'].to_list()
"""


# Active in model function definition
def active_in_model(twin_code_list, type_code_col, active_in_model_col):
    if type_code_col in twin_code_list:
        active_in_model_col = 'Yes'
    else:
        active_in_model_col = 'No'
    return active_in_model_col

# Active in model function definition
def scheduled_in_amr(amr_code_list, type_code_col, scheduled_in_amr_col):
    if type_code_col in amr_code_list:
        scheduled_in_amr_col = 'Yes'
    else:
        scheduled_in_amr_col = 'No'
    return scheduled_in_amr_col

wb = wb = load_workbook(filename=ir_file, read_only=False, keep_vba=False, data_only=True, keep_links=False)


for i in range(0, len(ir_sheets)):
    df = pd.read_excel(ir_file, ir_sheets[i], header=2)
    try:
        df['Active in Model'] = df.apply(lambda row: active_in_model(active_codes_list, row['Code (MM_Type)'], row['Active in Model']), axis=1)
        #print(df.columns)
    except KeyError:
        #print(df.columns)
        df['Active in Model'] = df.apply(lambda row: active_in_model(active_codes_list, row['Asset Type (MM_Type)'], row['Active in Model']), axis=1)
    
    df.to_excel(Path(output_path, ir_sheets[i]).with_suffix('.xlsx'), ir_sheets[i])
    ws = wb[ir_sheets[i]]
"""     for tbl in ws.tables.values():
        for row in ws[tbl.ref]:
            #for cell in row:
                #print(cell.column)
            print(tbl.name, tbl.ref) """
    #print(ws.tables.values())

amr_codes = [] #Define amr code list

data_dir = Path(ir_file).parent # return data directory

"""
Iterate over data directory and process amr registers for codes
"""
for dir in data_dir.iterdir():
    if dir.is_file() and 'REG-XIM-NAP' in dir.name:
        wb = load_workbook(dir)
        
        try:
            ws = wb['Validation Tables']
            amr_codes_table = ws.tables['Equipment_List']
            am_cod_ref = amr_codes_table.ref
            pattern = '[0-9]{1,}'
            start_row = int(re.findall(pattern, am_cod_ref)[0])
            end_row = int(re.findall(pattern, am_cod_ref)[1])

            """
            Iterate over excel worksheet using table reference form above
            """
            for row in ws.iter_rows(min_row=start_row, max_row=end_row, max_col=4):
                if 'Asset Hive' not in  row[2].value:
                  amr_codes.append(row[0].value)  
                #print(row[0].value, row[2].value)
            
        except:
            continue       
        #print(dir.name, dir)
print(len(amr_codes))
amr_codes = list(dict.fromkeys(amr_codes))
print(amr_codes.count('AAV'))

""" #Processing HWW Type codes
hww = pd.read_excel(ir_path, aud_ir_sheetname, header=2)
hww['Active in Model'] = hww.apply(lambda row: active_in_model(arc_active_codes_list, row['Code (MM_Type)'], row['Active in Model']), axis=1)
print(hww.columns)

# Processing RSHP Type Codes
rshp = pd.read_excel(ir_path, rshp_ir_sheetname, header=2) 
rshp['Active in Model'] = rshp.apply(lambda row: active_in_model(arc_active_codes_list, row['Code (MM_Type)'], row['Active in Model']), axis=1)

# Processing AUD Type Codes
aud = pd.read_excel(ir_path, aud_ir_sheetname, header=2) 
aud['Active in Model'] = aud.apply(lambda row: active_in_model(aud_active_codes_list, row['Code (MM_Type)'], row['Active in Model']), axis=1)



# hww['Active in Model'] = hww['Code (MM_Type)'].apply(lambda x: 'Yes' if x in arc_active_codes_list else 'No')

hww.to_excel(hww_out_path, 'output')
rshp.to_excel(rshp_out_path, 'output')
aud.to_excel(aud_out_path, 'output') """
