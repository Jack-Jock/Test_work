from openpyxl import load_workbook
import requests
import csv

"""
Ускладніть програму з першого завдання наступним фільтром 2. З посадою керівника Ректор
"""

print("Коди регіонів: 01, 05, 07, 12, 14, 18, 21, 23, 26, 32, 35, 44, 46, 48, 51, 53, 56, 59, 61, 63, 65, 68, 71, 73, 74, 80, 85")
num = str(input("Введіть код регіону: "))

wb = load_workbook('regions.xlsx')
sheet = wb['Аркуш1']
column = sheet['A']
code_index = []

for i in range(len(column)):
    cod = column[i].value
    code_index.append(cod)
code_index.pop(0)

if num in code_index:
    r = requests.get(f'https://registry.edbo.gov.ua/api/universities/?ut=1&lc={num}&exp=json')
    universities: list = r.json()

    filtered_data = [{num: row[num] for num in ['university_name', 'university_address', 'university_director_post', 'university_director_fio']}
                     for num in ['university_director_post'] for row in universities if row[num] == "Ректор"
                     or row[num] == "ректор"]

    with open('address_rector.csv', mode='w', encoding='cp1251', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=filtered_data[0].keys())
        writer.writeheader()
        writer.writerows(filtered_data)

    print("Данні записані у файли address_rector.csv!")

else:
    print("Помилка: такого регіона не існіє")