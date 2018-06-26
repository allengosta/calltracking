import requests as req
import xlsxwriter as xls

req.get("https://calltracking.ru/api/login.php?account_type=calltracking&login=oksana@soldco.ru&password=CiM6cd34c88YtU"
        "&service=analytics")

response = req.get("https://calltracking.ru/api/get_data.php?project=5056,5619,5427,5425,5426,6304"
                   "&auth=calltracking-743d58c4db514421cd179921e7662a09-4507-c75a9ab839b47fcb64719670cda9a9a849eaf8c2"
                   "&dimensions=ct%3Adatetime,ct%3Acaller,ct%3Asource,ct%3Adsource,ct%3Admedium,ct%3Adcampaign"
                   "&metrics=ct%3Acalls&sort=-ct%3Acalls &start-date=2018-01-01"
                   "&end-date=2018-07-01&max-results=9999&start-index=0&")

print(response.raw)
print(response.json())
js = response.json()
    #json.loads(response.text)
#print(js['data'])

xlsx = xls.Workbook("example.xlsx")
wk = xlsx.add_worksheet()

wk.set_column("A:H", 22)

bold = xlsx.add_format({'bold': True})

wk.write(0, 0, "Дата контакта", bold)
wk.write(0, 1, "Номер контакта", bold)
wk.write(0, 2, "Рекламный источник", bold)
wk.write(0, 3, "utm_source", bold)
wk.write(0, 4, "utm_medium", bold)
wk.write(0, 5, "utm_campaign", bold)
wk.write(0, 6, "предпочтения клиента", bold)
wk.write(0, 7, "проект", bold)

row = 1
col = 0

for i in js['data']:
    for j in js['data'][i]:
        for k in js['data'][i][j]:
            for z in js['data'][i][j][k]:
                for z2 in js['data'][i][j][k][z]:
                    for z3 in js['data'][i][j][k][z][z2]:
                            wk.write(row, col, i)
                            wk.write(row, col + 1, j)
                            wk.write(row, col + 2, k)
                            wk.write(row, col + 3, z)
                            wk.write(row, col + 4, z2)
                            wk.write(row, col + 5, z3)
                            row += 1

xlsx.close()
