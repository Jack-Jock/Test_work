from openpyxl import load_workbook
import requests
import csv

"""
Ускладніть програму з другого завдання можливістю фільтрування за будь-яким з наявних значень поля
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

    print(" Оберіть варіант за яким ви бажаете відфільтрувати поле посади керівника: \n\n   "
          "1 - Університети без фільтру посад. \n   2 - Університети в яких ніхто не керує. \n   "
          "3 - Університети керівну посаду в яких займає Ректор. \n   "
          "4 - Університети керівну посаду в яких займає Директор. \n   "
          "5 - Університети в яких інша особа виконує обов'язки ректора. \n   "
          "6 - Університети в яких інша особа виконує обов'язки директора. \n   "
          "7 - Університети керівну посаду в яких займає Ректор або Директор. \n   "
          "8 - Університети в яких інша особа виконує обов'язки ректора або директора. \n   ")

    choice = int(input("Ваш вибір - "))

    if choice == 1:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities]

    elif choice == 2:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == ""]

    elif choice == 3:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "Ректор"
                     or row[num] == "ректор"]

    elif choice == 4:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "Директор"
                     or row[num] == "директор"]

    elif choice == 5:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "В.о. ректора"]

    elif choice == 6:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "В.о. директора"]

    elif choice == 7:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "Ректор"
                     or row[num] == "ректор" or row[num] == "Директор"
                     or row[num] == "директор"]

    elif choice == 8:
        filtered_data = [{num: row[num] for num in
                          ['university_name', 'university_address', 'university_director_post',
                           'university_director_fio']}
                         for num in ['university_director_post'] for row in universities if row[num] == "В.о. ректора"
                     or row[num] == "В.о. директора"]

    else:
        print("Помилка: введіть номер варіанту, який є в списку")
        exit()



    with open('address_director_post_filter.csv', mode='w', encoding='cp1251', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=filtered_data[0].keys())
        writer.writeheader()
        writer.writerows(filtered_data)

    print("Данні записані у файли address_director_post_filter.csv!")

else:
    print("Помилка: такого регіона не існіє")