from openpyxl import Workbook

wb = Workbook()

ws = wb.active
ws.title = "Programa"

wb.save("excel.xlsx")

print("Tu archivo excel fue creado con exito")
