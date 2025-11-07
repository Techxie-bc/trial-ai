import csv
import os
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

# Mapping of LGA names to their corresponding CSV file names
csv_file_mapping = {
    "Aboh Mbaise": "aboh_mbaise.csv",
    "Ahiazu Mbaise": "ahiazu_mbaise.csv",
    "Ehime Mbano": "ehime_mbano.csv",
    "Ezinihitte Mbaise": "ezinihitte.csv",
    "Ideato North": "ideato_north.csv",
    "Ideato South": "ideato_south.csv",
    "Ihitte Uboma": "ihitte_uboma.csv",
    "Ikeduru": "ikeduru.csv",
    "Isiala Mbano": "isiala_mbano.csv",
    "Isu": "isu.csv",
    "Mbaitoli": "mbaitoli.csv",
    "Ngor Okpala": "ngor_okpala.csv",
    "Njaba": "njaba.csv",
    "Nkwerre": "nkwerre.csv",
    "Nwangele": "nwangele.csv",
    "Obowo": "obowo.csv",
    "Oguta": "oguta.csv",
    "Ohaji Egbema": "ohaji_egbema.csv",
    "Okigwe": "okigwe.csv",
    "Onuimo": "onuimo.csv",
    "Orlu": "orlu.csv",
    "Orsu": "orsu.csv",
    "Oru East": "oru_east.csv",
    "Oru West": "oru_west.csv",
    "Owerri Municipal": "owerri_municipal.csv",
    "Owerri North": "owerri_north.csv",
    "Owerri West": "owerri_west.csv"
}

# Create a new workbook
wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# Function to load CSV data
def load_csv_data(csv_filename):
    try:
        file_path = rf'C:\Users\USER\Desktop\trial ai\{csv_filename}'
        with open(file_path, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            data = list(csv_reader)
        return data
    except FileNotFoundError:
        print(f"CSV file not found: {csv_filename}")
        return None

# Create sheets for each LGA
for lga in imo_lgas:
    ws = wb.create_sheet(title=lga)
    
    # Check if we have CSV data for this LGA
    if lga in csv_file_mapping:
        csv_data = load_csv_data(csv_file_mapping[lga])
        if csv_data:
            for row in csv_data:
                ws.append(row)
            print(f"Added data to {lga} sheet ({len(csv_data)-1} projects)")
        else:
            # Add headers for empty sheets if CSV not found
            headers = ["S/N", "WARD", "PROJECT TO BE EXECUTED", "LOCATION"]
            ws.append(headers)
            print(f"Created empty sheet for {lga} (CSV not found)")
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
