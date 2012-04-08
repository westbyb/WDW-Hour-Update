#Author: Brian Westby
#File: parse.py
#Parses the hours of operation from calendar site.

import urllib
import re
import time
desmonth = raw_input("What month do you want to update (e.g. 04 for April, \
11 for November, etc)?: ")
readfile = raw_input("What is the name of the excel file you will be editing (include .xls)?: ")
outfile = raw_input("What do you want the updated file to be saved as (include .xls)?: ")
parkh = dict()
events = dict()

def main():
    parsehours()
    time.sleep(2)
    exceledit()

def parsehours():
    source = urllib.urlopen('http://disneyworld.disney.go.com/parks/magic-kingdom/calendar/')
    page = source.readlines()
    print 'Opening webpage...'
    curdate = 0
    #look for #:## AM - #:## PM
    date = r'2012'+desmonth+'\d{02}'
    hours = r'\d+:0{2}\s\w{2}\s-\s\d+:0{2}\s\w{2}'
    type = 'Park Hours|Extra Magic Hours'
    #go through page line by line
    for line in page:
        times = re.findall(hours, line.lower())
        types = re.findall(type, line)
        dates = re.search(date, line)
        #if date is found
        if dates:
            start = dates.start()
            end = dates.end()
            #get date with format on website
            rd = line[start:end]
            #if the month starts with a 0, strip the 0
            if rd[4:5] == '0':
                #if the day starts with a 0, strip the 0
                if rd[6:7] == '0':
                    curdate = rd[5:6] + "/" + rd[7:8] + "/" + rd[:4]
                else:
                    curdate = rd[5:6] + "/" + rd[6:8] +"/" + rd[:4]
            else:
                #if the day starts with a 0, strip the 0
                if rd[6:7] == '0':
                    curdate = rd[4:6] + "/" + rd[7:8] +"/" + rd[:4]
                else:
                    curdate = rd[4:6] + "/" + rd[6:8] + "/" + rd[:4]

        #if #:## - #:## is found, a date has been found
        if times:
            #create dictionary of dates, and within that dictionary, have dictionary of hour type and hours
            events = zip(types, times)
            parkh[curdate] = events
    print 'Data pulled from calendar'

def exceledit():
    #open excel sheet
    import xlrd, xlwt, xlutils
    import datetime
    from xlutils.copy import copy
    print 'Opening excel sheet...'
    book = xlrd.open_workbook(readfile, on_demand=True, formatting_info=True)
    print 'Creating and editing new excel sheet...'
    wbook = copy(book)
    wbook.dates_1904 = book.datemode
    print 'Done creating new excel sheet'
    
    sh = book.sheet_by_index(0)
    #iterate through dates in excel sheet
    for colnum in range(sh.ncols):
        date = sh.cell_value(3, colnum+4)
        #if xlrd finds a date
        if date:
            #grab date data
            year, month, day, hour, minute, second =  xlrd.xldate_as_tuple(date, book.datemode)
            format =  str(month) + "/" + str(day) + "/" + str(year)

            if month > int(desmonth):
                break

            #wbook.get_sheet(0).write(3, colnum+4, format)
            #if dates are within the month currently being edited
            if month == int(desmonth):
                #format excel date information to work with parkh dict
                #format =  str(month) + "/" + str(day) + "/" + str(year)
                print 'Editing ' + format
                #clear cells to eliminate old information
                wbook.get_sheet(0).write(6, colnum+4, "")
                wbook.get_sheet(0).write(5, colnum+4, "")
                wbook.get_sheet(0).write(7, colnum+4, "")
                #iterate through hour segments for that day
                for x in parkh[format]:
                    #if regular hours, insert in "HOURS" row
                    if x[0] == 'Park Hours':
                        wbook.get_sheet(0).write(6, colnum+4, x[1].replace(' ',''))
                    #if extra magic hours, insert in respective row
                    if x[0] == 'Extra Magic Hours':
                        #insert in morning row
                        if int(x[1][0:1]) in range(2,9):
                            wbook.get_sheet(0).write(5, colnum+4, x[1])
                        #insert in evening row
                        else:
                            wbook.get_sheet(0).write(7, colnum+4, x[1])

    print 'Done editing. Now saving...'
    wbook.save(outfile)
    print outfile+' saved'

#def adjustwait(colnum):
    

if __name__ == '__main__':
    main()
