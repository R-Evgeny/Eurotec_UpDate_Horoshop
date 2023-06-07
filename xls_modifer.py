import openpyxl
import csv
from tqdm import tqdm

#Формирование Renklod
book = openpyxl.open('renklod.xlsx', read_only=True)
sheet = book.active

with open(f"renklod.csv", "w", newline='', encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(
        (
            "Артикул",
            "Наличие"
        )
    )

for row in tqdm(range(2, sheet.max_row+1)):
    sku = sheet[row][5].value
    quantity = str(sheet[row][8].value)
    quantity = 'В наявності'
    with open(f"renklod.csv", "a", newline='', encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(
            (
                sku,
                quantity
            )
        )

print("Renklod - ОК")