import smartsheet
import datetime
import csv

column_map = {}

def get_cell_by_column_name(row, column_name):
    column_id = column_map[column_name]
    return row.get_column(column_id)

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = 'hfoz1pqg5qwcyxsq3u6bppr4y2'
sheetID = '6210591598634884'

# Initialize client
ss = smartsheet.Smartsheet(access_token)

eventsheet = ss.Sheets.get_sheet(sheetID, page_size=5000)

#print ("Loaded " + str(len(eventsheet.rows)) + " rows from sheet: " + eventsheet.name)

for column in eventsheet.columns:
    column_map[column.title] = column.id

today = datetime.date.today()
enddate = today + datetime.timedelta(days=14)
output = csv.writer (open("events.csv", "wb" ))

for row in eventsheet.rows:
    eventData = get_cell_by_column_name(row, "Date")

    if eventData.value:
        eventdate = datetime.datetime.strptime(eventData.value, '%Y-%m-%d').date()
    if today <= eventdate <= enddate:
        eventTitle = get_cell_by_column_name(row, "Event")
        eventLocation = get_cell_by_column_name(row, "City")
        eventDeliver = get_cell_by_column_name(row, "Live/Virtual")
        #print "---" + str(eventdate) + " - " + eventTitle.value + eventLocation.value + eventDeliver.value
        if (eventLocation.value=="Memphis" or eventLocation.value=="ALL") and ("Live" in eventDeliver.value):
            #print str(eventdate) + " - " + eventTitle.value
            output.writerow([str(eventdate),eventTitle.value])
