import os
import requests
import xlrd
import xlsxwriter
import time
import re
from tqdm import tqdm

# Lists
namelst = []
addresslst = []
citylst = []
statelst = []
postallst = []
countrylst = []
idlst = []
typelst = []

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

    # For every row in the document
    # tqdm gives us a loading bar
    for row in tqdm(range(1, sheet.nrows)):
        # This keeps us from overloading the Google Maps API
        time.sleep(.1)

        # Update our lists for writeOut()
        namelst.append(sheet.cell(row, nameCol).value)
        addresslst.append(sheet.cell(row, addressCol).value)
        citylst.append(sheet.cell(row, cityCol).value)
        statelst.append(sheet.cell(row, stateCol).value)
        postallst.append(sheet.cell(row, postalCol).value)
        countrylst.append(sheet.cell(row, countryCol).value)

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
            print('Could not find Google listing for: ' + sheet.cell(row, queryCol).value + '\n')
            idlst.append(0)
            typelst.append(0)

        # If we get a result
        else:
            idlst.append(data['results'][0]['place_id'])

            # I'm grabbing the storeType which is a list and making it into a space-separated string
            storeType = ' '.join(data['results'][0]['types'])
            # Estbalishment, POI, and store are all common types that aren't valuable to know
            storeType = storeType.replace('establishment', '')
            storeType = storeType.replace('point_of_interest', '')
            storeType = re.sub(r'\bstore\b', '', storeType)

            # I split the string back into a list and then rejoin it with commas to remove empty white space from the previous replaces
            storeType = ', '.join(storeType.split())
            storeType = storeType.replace('_', ' ')

            # Google doesn't keep track of Toy/Gift stores, but given all their other types, if something is blank it's very likely a toy/gift shop or something very rare
            if not storeType:
                storeType = 'toy/gift/misc (?)'

            # Add the type to the list
            typelst.append(storeType)

    # If we get an overloaded API error just redo it
    elif(data['status'] == 'UNKNOWN_ERROR'):
        scrape(query, region, key, sheet, row)

    # If there's some other error like REQUEST_DENIED
    else:
        print(data['status'] + ' error for: ' + sheet.cell(row, queryCol).value + '\n')
        idlst.append(0)
        typelst.append(0)

# Writes out the finished files
def writeOut():
    workbook = xlsxwriter.Workbook('../data/output.xlsx')
    worksheet = workbook.add_worksheet()
    row = 1

    headerFormat = workbook.add_format({'font_color': 'white', 'bg_color': '#4f81bd', 'align': 'center'})

    worksheet.write(0, 0, 'Name', headerFormat)
    worksheet.write(0, 1, 'Types', headerFormat)
    worksheet.write(0, 2, 'Address', headerFormat)
    worksheet.write(0, 3, 'City', headerFormat)
    worksheet.write(0, 4, 'State', headerFormat)
    worksheet.write(0, 5, 'ZIP', headerFormat)
    worksheet.write(0, 6, 'Country', headerFormat)
    worksheet.write(0, 7, 'ID', headerFormat)

    for i in range(0, len(namelst)):
        if(typelst[i] != 0):
            worksheet.write(row, 0, namelst[i])
            worksheet.write(row, 1, typelst[i])
            worksheet.write(row, 2, addresslst[i])
            worksheet.write(row, 3, citylst[i])
            worksheet.write(row, 4, statelst[i])
            worksheet.write(row, 5, postallst[i])
            worksheet.write(row, 6, countrylst[i])
            worksheet.write(row, 7, idlst[i])
            row = row + 1

if __name__ == '__main__':
    readIn()
    writeOut()
