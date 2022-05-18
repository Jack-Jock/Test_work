from openpyxl import load_workbook
import requests
import csv

"""
Для всіх
Запитати у користувача код регіону
Отримати ЗВО з вказаного користувачем регіону
Зберегти всі дані у файл universities.csv у форматі csv
Збережіть ті ж дані у файл universities_<код регіону>.csv, наприклад universities_80.csv
Якщо регіон не зі списку доступних, то повідомити про це користувачеві у консолі
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


    with open('universities.csv', mode='w', encoding='cp1251', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=universities[0].keys())
        writer.writeheader()
        writer.writerows(universities)

    with open(f'universities_{num}.csv', mode='w', encoding='cp1251', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=universities[0].keys())
        writer.writeheader()
        writer.writerows(universities)

    print(f"Данні записані у файли universities.csv и universities_{num}.csv!")

else:
    print("Помилка: такого регіона не існіє")

