import csv
import openpyxl
import os

# Wczytuję nazwy z exela , weź je z valve listy
excel_file = r'D:\CSV\Zawory.xlsx'

workbook = openpyxl.load_workbook(excel_file)
worksheet = workbook.active
names = [cell.value for cell in worksheet['A']]           # Wczytuję nazwę z  exela , weź je z valve listy
valvestime = [cell.value for cell in worksheet['B']]      # Wczytuję czas z exela , weź go z valve listy

# Na podstawie nazwy zrób osobny csv dla każdego zaworu
base_folder = r'D:\CSV\Valves'
#Zrób nowy folder "Valves"
os.makedirs(base_folder, exist_ok=True)

for valves, name in enumerate(names):
    csv_name = f'V{valves + 1}.csv'
    csv_file = os.path.join(base_folder, csv_name)
    csv_file_obj = open(csv_file, 'w', newline='')
    writer = csv.writer(csv_file_obj)

    writer.writerow([name + ';'])
    writer.writerow([str(int(int(valvestime[valves]) / 10)+5) + ';'])  # Dzię DN na /10 i dodaje 5 sekund dodatkowo
    #writer.writerow(["45;"])

    csv_file_obj.close()

print(f'Pliki z zaworami csv są tutaj: {base_folder}')