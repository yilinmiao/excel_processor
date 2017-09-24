from openpyxl import load_workbook

#init workbook
filename = '2174 Schedule Planning  - Instructor Analysis.xlsx'
wb = load_workbook(filename)
sheetnames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetnames[0])

#store schedule sheet to dict
schedule = {}
for row in range(1, ws.max_row + 1):
    temp_list = []
    index = row - 1
    col1 = ws.cell(row=row, column=1).value
    col2 = ws.cell(row=row, column=2).value
    col3 = ws.cell(row=row, column=3).value
    col4 = ws.cell(row=row, column=4).value
    col5 = ws.cell(row=row, column=5).value
    col6 = ws.cell(row=row, column=6).value
    col7 = ws.cell(row=row, column=7).value
    temp_list = [col1, col2, col3, col4, col5, col6, col7]
    schedule[index] = temp_list

#generate output
instructor_output = {}
instructor_output[0] = ['Instructor']
subject = ''
num = 0;
section = 0;
i = 0;
for row in range(1, ws.max_row):
    if ((schedule[row][0] is None) and (schedule[row][4] is None)):
        continue;
    if (schedule[row][1] is None):
        schedule[row][1] = subject
    else:
        subject = schedule[row][1]
    if (schedule[row][2] is None):
        schedule[row][2] = num
    else:
        num = schedule[row][2]
    if (schedule[row][3] is None):
        schedule[row][3] = section
    else:
        section = schedule[row][3]
    if (schedule[row][3] <= 9):
        schedule[row][3] = '0' + str(schedule[row][3])
    if (schedule[row][4] != 'TBA'):
        if (schedule[row][5] < 1000):
            schedule[row][5] = (str)(schedule[row][5])[0:1] + ':' + (str)(schedule[row][5])[1:]
        else:
            schedule[row][5] = (str)(schedule[row][5])[0:2] + ':' + (str)(schedule[row][5])[2:]
        if (schedule[row][6] < 1000):
            schedule[row][6] = (str)(schedule[row][6])[0:1] + ':' + (str)(schedule[row][6])[1:]
        else:
            schedule[row][6] = (str)(schedule[row][6])[0:2] + ':' + (str)(schedule[row][6])[2:]
        temp = schedule[row][1] + ' ' + (str)(schedule[row][2]) + '-' + (str)(schedule[row][3]) + ', ' + schedule[row][4] + ', ' + (str)(schedule[row][5]) + '-' + (str)(schedule[row][6])
    else:
        temp = schedule[row][1] + ' ' + (str)(schedule[row][2]) + '-' + (str)(schedule[row][3]) + ', ' + schedule[row][4]

    if ((schedule[row][0] is None)):
        instructor_output[i].append(temp)
    else:
        i += 1
        instructor_output[i] = [schedule[row][0], temp]

max_len = max([len(n) for n in instructor_output.values()])
for num in range (1, max_len):
    instructor_output[0].append('Class' + str(num))
temp = sheetnames[1]
ws_archived = wb.get_sheet_by_name(temp)
wb.remove_sheet(ws_archived)
ws_new = wb.create_sheet(temp)
for row in instructor_output:
    for col in range(0, len(instructor_output[row])):
        ws_new.cell(column=col + 1, row=row + 1, value=instructor_output[row][col])
wb.save(filename)



