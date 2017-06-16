import os
import requests
import xlrd
import xlsxwriter
import time
from tqdm import tqdm

# Lists
namelst = []
addresslst = []
idlst = []
typelst = []
cachelst = []
errors = []

# Columns
nameCol = 1
addressCol = 4
cityCol = 6
stateCol = 7
postalCol = 8
countryCol = 9
queryCol = 16

# Reads our file
def readIn():
    # On demand helps keep memory requirements low by only processing parts of the sheet ad hoc
    excel = xlrd.open_workbook('../data/locations.xlsx', on_demand = True)
    sheet = excel.sheet_by_index(0)
    process(sheet)

# Scrapes for lat/long coordinates
def process(sheet):
    key = input('OPTIONAL\nPlease enter your Google API key. Enter to skip.\n')

    # Tqdm gives us a loading bar
    for row in tqdm(range(1, sheet.nrows)):
        # This keeps us from overloading the Google Maps API
        time.sleep(.1)

        # Update our lists
        namelst.append(sheet.cell(row, nameCol).value)
        addresslst.append(sheet.cell(row, queryCol).value)

        # Google considers Guam its own country, so this fixes that case for region-biasing
        if(sheet.cell(row, stateCol).value == 'GU'):
            region = 'GU'

        else:
            region = sheet.cell(row, countryCol).value

        # Grab the lat/long coordinats
        scrape(sheet.cell(row, queryCol).value, region, key, sheet, row)

def scrape(query, region, key, sheet, row):
    # This is the Google Maps API Query URL format
    response = requests.get('https://maps.googleapis.com/maps/api/geocode/json?key=' + key + '&components=country:' + region + '&address=' + query)

    # Grab the JSON response
    data = response.json()

    # If the API has no error
    if(data['status'] == 'OK'):
        # Region biasing seems to remove the ZERO_RESULTS error and just give the lat/long coords for the center of the United States or Canada. This catches that.
        if(data['results'][0]['formatted_address'] == 'United States' or data['results'][0]['formatted_address'] == 'Canada'):
            errors.append('Could not find Google listing for: ' + sheet.cell(row, queryCol).value)
            idlst.append(0)
            typelst.append(0)

        # If we get a result
        else:
            idlst.append(data['results'][0]['place_id'])

            storeType = [x for x in data['results'][0]['types'] if x not in ('establishment', 'point_of_interest', 'store')]

            # Google doesn't keep track of Toy/Gift stores, but given all their other types, if something is blank it's very likely a toy/gift shop or something very rare
            if not storeType:
                storeType.append('toy_gift_misc')

            # Add the type to the list
            typelst.append(storeType)

    # If we get an overloaded API error just redo it
    elif(data['status'] == 'UNKNOWN_ERROR'):
        scrape(query, region, key, sheet, row)

    # If there's some other error like REQUEST_DENIED
    else:
        errors.append(data['status'] + ' error for: ' + sheet.cell(row, queryCol).value)
        idlst.append(0)
        typelst.append(0)

# Writes out the finished files
def output():
    workbook = xlsxwriter.Workbook('../data/output.xlsx')
    worksheet = workbook.add_worksheet()
    row = 1

    headerFormat = workbook.add_format({'font_color': 'white', 'bg_color': '#4f81bd', 'align': 'center'})

    # Headers
    worksheet.write(0, 0, 'Name', headerFormat)
    worksheet.write(0, 1, 'Type(s)', headerFormat)
    worksheet.write(0, 2, 'Address', headerFormat)
    worksheet.write(0, 3, 'ID', headerFormat)

    # Data
    for i in range(0, len(namelst)):
        if(typelst[i] != 0):
            worksheet.write(row, 0, namelst[i])
            worksheet.write(row, 1, ', '.join(typelst[i]))
            worksheet.write(row, 2, addresslst[i])
            worksheet.write(row, 3, idlst[i])
            row = row + 1

def errorReport():
    if errors:
        with open('../data/errors.txt', 'w') as e:
            for error in errors:
                e.write(error + '\n')

if __name__ == '__main__':
    readIn()
    output()
    errorReport()
