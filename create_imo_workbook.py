import csv
from openpyxl import Workbook

# List of all 27 LGAs in Imo State (cleaned for Excel sheet names)
imo_lgas = [
    "Aboh Mbaise",
    "Ahiazu Mbaise", 
    "Ehime Mbano",
    "Ezinihitte Mbaise",
    "Ideato North",
    "Ideato South",
    "Ihitte Uboma",  # Removed /
    "Ikeduru",
    "Isiala Mbano",
    "Isu",
    "Mbaitoli",
    "Ngor Okpala",
    "Njaba",
    "Nkwerre",
    "Nwangele",
    "Obowo",
    "Oguta",
    "Ohaji Egbema",  # Removed /
    "Okigwe",
    "Onuimo",
    "Orlu",
    "Orsu",
    "Oru East",
    "Oru West",
    "Owerri Municipal",
    "Owerri North",
    "Owerri West"
]

# Create a new workbook
wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# Read the existing Ezinihitte Mbaise data
ezinihitte_data = []
try:
    with open(r'C:\Users\USER\Desktop\trial ai\ezinihitte.csv', 'r', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        ezinihitte_data = list(csv_reader)
    print(f"Loaded {len(ezinihitte_data)} rows from CSV file")
except FileNotFoundError:
    print("CSV file not found. Creating empty sheets.")

# Create sheets for each LGA
for lga in imo_lgas:
    ws = wb.create_sheet(title=lga)
    
    # If this is Ezinihitte Mbaise, add the existing data
    if lga == "Ezinihitte Mbaise" and ezinihitte_data:
        for row in ezinihitte_data:
            ws.append(row)
        print(f"Added data to {lga} sheet ({len(ezinihitte_data)-1} projects)")
    else:
        # Add headers for empty sheets
        headers = ["S/N", "WARD", "PROJECT TO BE EXECUTED", "LOCATION"]
        ws.append(headers)
        print(f"Created empty sheet for {lga}")

# Save the workbook
excel_file = r'C:\Users\USER\Desktop\trial ai\Imo_State_Ward_Projects.xlsx'
wb.save(excel_file)

print(f"\nSuccessfully created Excel workbook: {excel_file}")
print(f"Total sheets created: {len(imo_lgas)}")
print("\nAll 27 LGA sheets are ready for ward project data entry!")
