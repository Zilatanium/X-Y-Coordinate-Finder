import openpyxl
import requests

def geocode(address):
    api_key = ''  # Replace with your Google Maps Geocoding API key
    base_url = 'https://maps.googleapis.com/maps/api/geocode/json?'

    # Construct the request URL
    url = base_url + f'address={address}&key={api_key}'

    # Send the HTTP GET request
    response = requests.get(url)
    data = response.json()

    # Parse the response and extract the coordinates
    if data['status'] == 'OK':
        result = data['results'][0]
        location = result['geometry']['location']
        latitude = location['lat']
        longitude = location['lng']
        return latitude, longitude
    else:
        return None

# Load the Excel file
file_path = ''  # Replace with the actual file path
wb = openpyxl.load_workbook(file_path)
sheet = wb['Sheet1']  # Replace with the actual sheet name

# Iterate over rows and geocode addresses
for row in range(2, sheet.max_row + 1):
    address = sheet[f'C{row}'].value
    town = sheet[f'D{row}'].value
    state = sheet[f'E{row}'].value

    full_address = f'{address}, {town}, {state}'
    coordinates = geocode(full_address)

    if coordinates:
        latitude, longitude = coordinates
        sheet[f'V{row}'].value = f'{latitude}, {longitude}'
    else:
        sheet[f'V{row}'].value = 'Not found'

# Save the modified Excel file
wb.save('')  # Replace with the desired output file path
print('Finish')