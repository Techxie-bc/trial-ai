import csv
from openpyxl import Workbook

# Input and output file paths
input_csv = r'C:\Users\USER\Desktop\trial ai\LGA_Ward_Only.csv'
output_xlsx = r'C:\Users\USER\Desktop\trial ai\LGA_Ward_Only.xlsx'

# Create a new workbook
wb = Workbook()
ws = wb.active
ws.title = "LGA_Ward_Mapping"

# Read CSV and write to Excel
with open(input_csv, 'r', encoding='utf-8') as csvfile:
    reader = csv.reader(csvfile)
    
    row_count = 0
    for row in reader:
        row_count += 1
        # Write each row to the Excel worksheet
        for col_num, cell_value in enumerate(row, 1):
            ws.cell(row=row_count, column=col_num, value=cell_value)

# Save the Excel file
wb.save(output_xlsx)

print(f"Successfully converted CSV to Excel: {output_xlsx}")
print(f"Total rows processed: {row_count}")
print(f"Total columns: {len(row) if 'row' in locals() else 0}")

wb.close()
