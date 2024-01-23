import pandas as pd
import requests
import lxml
from bs4 import BeautifulSoup

# Read the Excel file
excel_path = 'C:\\path\\to\\excel\\Final_Filtered_Data.xlsx'
df = pd.read_excel(excel_path, usecols=[0])

# Initialize a dictionary to store the results
emdb_resolutions = {}

# Iterate over the codes
for code in df.iloc[:, 0]:
    url = f"https://www.ebi.ac.uk/emdb/{code}"
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'lxml')
        entry_box = soup.find('div', class_='entry_header_container')

        if entry_box:
            resolution_tag = entry_box.find('b', string=lambda t: t and any(char.isdigit() for char in t))
            if resolution_tag:
                resolution = resolution_tag.get_text(strip=True)
                print(code," ", resolution)
                emdb_resolutions[code] = resolution
            else:
                print(f"Resolution not found for {code}")
        else:
            print(f"Entry box not found for {code}")
    else:
        print(f"Failed to retrieve data for {code}")

# Print the dictionary
print(emdb_resolutions)

# Generate excel file
resolutions_df = pd.DataFrame(list(emdb_resolutions.items()), columns=['emdb_id', 'resolution'])
output_path = 'C:\\path\\to\\excel\\EMDB_Resolutions.xlsx'
resolutions_df.to_excel(output_path, index=False)

print(f"Excel file generated at {output_path}")