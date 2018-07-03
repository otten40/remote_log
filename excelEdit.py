import os
import openpyxl

os.chdir('C:\\Users\\320051\\Desktop')

remote_log = openpyxl.load_workbook('Remote Access Log copy.xlsx')

sheet = remote_log['Sheet1']
past_logs = {}

for row in range(2, sheet.max_row + 1):
    remote_user: object = sheet['B' + str(row)].value
    print(remote_user)

