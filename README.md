# Imo State Ward Projects Data Management

This project manages ward project data for all 27 Local Government Areas (LGAs) in Imo State, Nigeria.

## Project Structure

```
trial ai/
├── README.md                    # This documentation
├── create_imo_workbook.py      # Main script to generate Excel workbook
├── ezinihitte.csv              # Sample CSV data for Ezinihitte Mbaise LGA
└── Imo_State_Ward_Projects.xlsx # Generated Excel workbook (after running script)
```

## Overview

The project converts ward project data from CSV format into a comprehensive Excel workbook with separate sheets for each of the 27 LGAs in Imo State.

## Requirements

- Python 3.x
- openpyxl library

### Installing Dependencies

```bash
pip install openpyxl
```

## Data Format

Each CSV file should follow this structure:

| Column | Description |
|--------|-------------|
| S/N | Serial Number |
| WARD | Ward name and number |
| PROJECT TO BE EXECUTED | Description of the project |
| LOCATION | Specific location within the ward |

### Example CSV Format:
```csv
S/N,WARD,PROJECT TO BE EXECUTED,LOCATION
1,Amumara Ward 1,Proposal for Building of Lockup Shops (5),Nkwo-Otulu Aumuma Market Square
2,Itu Ward 2,Construction of a 3 Room Administrative Block,Itu Secondary School
```

## Usage

### 1. Generating the Excel Workbook

Run the main script to create an Excel workbook with all 27 LGA sheets:

```bash
python create_imo_workbook.py
```

This will:
- Create `Imo_State_Ward_Projects.xlsx` with 27 sheets
- Each sheet is named after an LGA
- Automatically populate the "Ezinihitte Mbaise" sheet with existing CSV data
- Create empty sheets with headers for other LGAs

### 2. Adding Data for Other LGAs

To add data for other LGAs:

1. **Create CSV files** for each LGA following the format above
2. **Name the CSV files** using the LGA name (e.g., `aboh_mbaise.csv`)
3. **Modify the script** to include additional LGA data loading

### 3. Data Entry Process

#### From Images/Documents:
1. Extract text data from images/documents
2. Convert to CSV format with proper headers
3. Save as `[lga_name].csv`
4. Run the script to update the Excel workbook

#### Direct Excel Entry:
1. Open the generated `Imo_State_Ward_Projects.xlsx`
2. Navigate to the appropriate LGA sheet
3. Enter data directly following the column structure

## Complete List of Imo State LGAs

The workbook includes sheets for all 27 LGAs:

1. Aboh Mbaise
2. Ahiazu Mbaise
3. Ehime Mbano
4. Ezinihitte Mbaise
5. Ideato North
6. Ideato South
7. Ihitte Uboma
8. Ikeduru
9. Isiala Mbano
10. Isu
11. Mbaitoli
12. Ngor Okpala
13. Njaba
14. Nkwerre
15. Nwangele
16. Obowo
17. Oguta
18. Ohaji Egbema
19. Okigwe
20. Onuimo
21. Orlu
22. Orsu
23. Oru East
24. Oru West
25. Owerri Municipal
26. Owerri North
27. Owerri West

## Script Modifications

### Adding More LGA Data

To modify the script for additional LGA data:

```python
# Add more CSV file reading logic
try:
    with open(f'C:\\Users\\USER\\Desktop\\trial ai\\{lga.lower().replace(" ", "_")}.csv', 'r', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        lga_data = list(csv_reader)
except FileNotFoundError:
    lga_data = None

# In the sheet creation loop
if lga_data:
    for row in lga_data:
        ws.append(row)
```

## Notes

- Excel sheet names cannot contain certain characters (`/`, `\`, `?`, `*`, `[`, `]`)
- The script automatically handles this by replacing problematic characters
- CSV files should be UTF-8 encoded to handle special characters properly
- Always backup your data before running scripts

## Troubleshooting

### Common Issues:

1. **ModuleNotFoundError: No module named 'openpyxl'**
   - Solution: Install openpyxl using `pip install openpyxl`

2. **ValueError: Invalid character found in sheet title**
   - Solution: The script automatically handles this, but ensure LGA names don't contain `/`, `\`, etc.

3. **FileNotFoundError for CSV files**
   - Solution: Ensure CSV files exist in the correct directory with proper naming

## Contributing

When adding new features or LGA data:

1. Follow the existing CSV format
2. Test the script with sample data
3. Update this README with any new processes
4. Ensure all 27 LGAs are properly handled

## Contact

For questions or issues with the data management process, refer to this documentation or contact the project maintainer.
