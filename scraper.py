
import scraperwiki
import requests
import xlrd
import datetime
import re
import urllib
import urllib2
import simplejson

def cellval(cell, datemode):
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            datetuple = xlrd.xldate_as_tuple(cell.value, datemode)
        except Exception, e:
            print "BAD", cell, e
            return str(cell)
        try:
            if datetuple[3:] == (0, 0, 0):
                return datetime.date(datetuple[0], datetuple[1], datetuple[2])
            return datetime.datetime(datetuple[0], datetuple[1], datetuple[2], datetuple[3], datetuple[4], datetuple[5])
        except ValueError, e:
            print "BAD value", datetuple, cell, e
            return str(cell)
    if cell.ctype == xlrd.XL_CELL_EMPTY:    return None
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:  return cell.value == 1
    return cell.value

def getCrimeData(year, month, url) :

    # The URL for a month's data
    url = "http://www.minneapolismn.gov/www/groups/public/@mpd/documents/webcontent/" + url
    # "wcms1p-104213.xlsx"
  
    # Get the data from the spreadsheet
    xlbin = scraperwiki.scrape(url)
    book = xlrd.open_workbook(file_contents=xlbin)
    
    # There should be one sheet that we care about called INCIDENTS
    sheet = book.sheets()[0]
 
    # Get the column headers and turn them into valid key names by removing periods
    keys = sheet.row_values(0)
    for col in range(0, len(keys)) :
        keys[col] = keys[col].replace('.', '')
        keys[col] = keys[col].replace(' ', '_')
        keys[col] = keys[col].replace('#', '')
    
    #For each row
    recordCount = 0
    print "Reading " + str(sheet.nrows -1) + " rows" 
    for rownumber in range(1, sheet.nrows):
        try:
            # create dictionary of the row values
            values = [ cellval(c, book.datemode) for c in sheet.row(rownumber) ]
            data = dict(zip(keys, values))
 
            data['YEAR'] = year
            data['MONTH'] = month        
            # Now that we have the clean address, geocode it
            # geocodeData(data)
                
            # Save th data to the ScraperWiki data store
            scraperwiki.sqlite.save(unique_keys=['YEAR', 'MONTH', 'NEIGHBORHOOD'], data=data)
            recordCount = recordCount +1
        except Exception as inst:
            print type(inst)     # the exception instance
            print inst.args      # arguments stored in .args
            print inst           # __str__ allows args to printed directly
            print "Failed to save row " + str(rownumber) + " of " + url
    print "Read " + str(recordCount)  + " incidents from " + url

def getFileList() :
    return [[2013,1, "wcms1p-104213.xlsx"],
    [2013,2, "wcms1p-105285.xlsx"],
    [2013,3, "wcms1p-106697.xlsx"],
    [2013,4, "wcms1p-108036.xlsx"],
    [2013,5, "wcms1p-109747.xlsx"],
    [2013,6, "wcms1p-110838.xlsx"],
    [2013,7, "wcms1p-112499.xlsx"],
    [2013,8, "wcms1p-114158.xlsx"],
    [2013,9, "wcms1p-115894.xlsx"]
    ]
    
def main () :
    files = getFileList()
    for row in  files :
        year = row[0]
        month = row[1]
        URI = row[2]
        print "Getting data for " + str(year) + " " + str(month) + " from " + URI
        getCrimeData(year, month, URI)

main()
# Saving data:
# unique_keys = [ 'id' ]
# data = { 'id':12, 'name':'violet', 'age':7 }
# scraperwiki.sql.save(unique_keys, data)


