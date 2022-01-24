current_month = '2021-12-31'
import pandas as pd
from _functions import *
from openpyxl import Workbook, load_workbook
wb = load_workbook('execute/GAAPIS.xlsx')
ws = wb.active

interest_income = '10'
interest_expense = '11'
gnii = '12'
pcl_income = '13'
nii = '14'
insurance = '17'
licensing_income = '18'
other_ = '19'
total_other = '20'
nii_net = '24'
sb = '27'
ooe = '28'
rc = '29'
toe = '30'
oibit = '32'
pfit = '34'
net_income = '36'

dates = pd.Series(['2021-01-31', '2021-02-28', '2021-03-31','2021-04-30','2021-05-31','2021-06-30','2021-07-31','2021-08-31','2021-09-30','2021-10-31','2021-11-30','2021-12-31'])
idx = pd.Series(['D','E','F','G','H','I','J','K','L','M','N','O'])
idx.index = dates 
print(idx) 


income_statement()
wb.save(f'execute/{current_month} GAAP IS.xlsx')