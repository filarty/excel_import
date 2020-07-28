import xlrd
"<--Импортируем лист из excel -->"

workbook = xlrd.open_workbook("D.xlsx")



wk = workbook.sheet_by_index(0)


list_colum = []
for j in range(0, 7):
        list_colum.append(wk.cell_value(1,j))

str_colum = list_colum[0]

first_name = str_colum.find(' ')

last_name = str_colum.rfind(' ')

mname = str_colum.rfind(" ")

print(str_colum[0:first_name])
print(str_colum[first_name:last_name])
print(str_colum[last_name:])