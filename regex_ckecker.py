import pandas as pd
import openpyxl
import re
from pathlib import Path
from xlrd import XLRDError

pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 100)
pd.set_option('display.width', 5000)

path = Path('./pliki/').absolute()
files = [str(p) for p in path.rglob('*.xlsx')]

result_plc_new = pd.DataFrame(columns=['Description', 'Files', 'RegexPass']).set_index('Description')

for file in files:
    print(file)

    try:
        workbook = openpyxl.load_workbook(file, data_only=True, read_only=True)
    except (EOFError, XLRDError):
        pass
    else:

        # PLC nowe
        sheet = workbook["Schränke,UV"]
        
        str_control_cabinet = ["BVS-Schrank", "OC.*CPC-24HP", "OVC-panel"]
        str_power_cabinet = ["Einspeiseschrank mit", "Einspeisung mit", "power cabinet with", "Power cabinet \(CE\)"]
        
        df = pd.DataFrame(sheet.values).iloc[:, 1:3].fillna(value=0).astype(str)

        for x in df.itertuples():
            try:
                number = int(x[2])
            except:
                continue
            else:
                if number > 0:
                    description = str(x[1])
                    if description not in result_plc_new.index:
                        pattern = "(" + ")|(".join(str_control_cabinet) + ")|(" + ")|(".join(str_power_cabinet) + ")"
                        result_plc_new.loc[description] = [
                            file,
                            bool(re.match(pattern, description, re.IGNORECASE)),
                        ]
                    else:
                        result_plc_new.loc[description, 'Files'] += ("\r\n" + str(file))

        workbook.close()

print(result_plc_new)

