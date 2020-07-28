import requests
import xlrd

esistens  = 0

button = str(input("Введите путь до файла  "))

workbook = xlrd.open_workbook(button)

url = "http://192.168.1.135:8080/mcsapi/api/v1/contacts"

headers = {

    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX2lkIjoiNWYxN2M4ZjA2NzViY2Q1ODAwNTI1OWM3IiwidXNlcm5hbWUiOiJhZG1pbiIsImlzX25hc191c2VyIjp0cnVlLCJpc190dXRvcmlhbF9kaXNwbGF5ZWQiOnRydWUsInNpZCI6InFnbjBhcnllIiwiaWF0IjoxNTk1ODUxNzM2LCJleHAiOjE2ODIyNTE3MzZ9.KrwOBO64YEE10lal5DPbzhQ_vuqLtwT1HAVlA31FkPI"
}


wk = workbook.sheet_by_index(0)
for i in range(1,wk.nrows):
    list_colum = []
    for j in range(0, 7):
        list_colum.append(wk.cell_value(i, j))

    params = {
          "fname": list_colum[1],
          "lname": list_colum[2],
          "mname": list_colum[0],
          "company_name": list_colum[3],
          "phones": [{"label": "MOBILE", "value": list_colum[4]}, {"label": "OFFICE", "value": list_colum[5]}],
          "im": [{"label": ".", "value": list_colum[6]}]
               }

    rec = requests.post(url, headers=headers, json=params)
    esistens += 1

    if rec.text:
         print("Готово",params["mname"])


print("Всего добавлено {} записей".format(esistens))