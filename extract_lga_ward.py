import csv

# Input and output file paths
input_file = r'C:\Users\USER\Desktop\trial ai\Imo_State_Ward_Projects_Consolidated.csv'
output_file = r'C:\Users\USER\Desktop\trial ai\LGA_Ward_Only.csv'

# Read the input file and extract only LGA and Ward columns
with open(input_file, 'r', encoding='utf-8') as infile:
    reader = csv.reader(infile)
    
    # Read header row
    header = next(reader)
    
    # Find the indices of LGA and WARD columns
    lga_index = header.index('LGA')
    ward_index = header.index('WARD')
    
    # Prepare output data
    output_data = []
    output_data.append(['LGA', 'WARD'])  # New header
    
    # Process each row
    for row in reader:
        if len(row) > max(lga_index, ward_index):
            lga = row[lga_index].strip()
            ward = row[ward_index].strip()
            
            # Skip empty rows
            if lga and ward:
                output_data.append([lga, ward])

# Write to output file
with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerows(output_data)

print(f"Successfully extracted LGA and Ward columns to: {output_file}")
print(f"Total rows processed: {len(output_data) - 1}")  # Subtract 1 for header

# Show first few rows
print("\nFirst 10 rows of the extracted data:")
for i, row in enumerate(output_data[:11]):
    print(f"Row {i}: {row}")
