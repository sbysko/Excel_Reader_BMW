
import pandas as pd
import openpyxl
from pathlib import Path
from xlrd import XLRDError

pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 100)
pd.set_option('display.width', 5000)

path = Path('./pliki/').absolute()
files = [str(p) for p in path.rglob('*.xlsx')]

workbook = openpyxl.load_workbook(files[0], data_only=True)

sheet = workbook["Sonstiges"]
df = pd.DataFrame(sheet.values).iloc[:, 1:3].fillna(value=0).astype(str)
str_vibn_step_2 = "Stufe 2"
str_vibn_step_3 = "Stufe 3"

str_vibn_step_2 = str_vibn_step_2.lower()
str_vibn_step_3 = str_vibn_step_3.lower()
df = df.applymap(str.lower)

vibn_step_2 = (df.loc[df[1].str.contains(str_vibn_step_2), [2]] != 'nein').sum().sum()
vibn_step_3 = (df.loc[df[1].str.contains(str_vibn_step_3), [2]] != 'nein').sum().sum()

print(vibn_step_2)
print(vibn_step_3)