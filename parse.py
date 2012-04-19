#Author: Brian Westby
#File: parse.py
#Parses the hours of operation from calendar site.

import urllib
import re
import time
from datetime import datetime, timedelta
from BeautifulSoup import BeautifulSoup
import xlrd, xlwt, xlutils

#readfile = raw_input("What is the name of the excel file you will be editing (include .xls)?: ")
#outfile = raw_input("What do you want the updated file to be saved as (include .xls)?: ")
parkh = dict()

def main():
    desmonth = '05'
    readfile = 'test.xls'
    outfile = 'new.xls'
    beausoupparse(desmonth)
    #desmonth = raw_input("What month do you want to update (e.g. 04 for April, 11 for November, etc)?: ")
    #readfile = raw_input("What is the name of the excel file you will be editing (include .xls)?: ")
    #outfile = raw_input("What do you want the updated file to be saved as (include.xls)?: ")
    time.sleep(2)
    exceledit(desmonth,readfile,outfile)

def formatdate(rd):
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
    return curdate

def parsetime(day,s):
    #parse 12-hour format
    return datetime.strptime(day+" "+s, '%m/%d/%Y %I:%M %p')

def beausoupparse(desmonth):
    #open the webpage and create a BeautifulSoup object with it
    print 'Opening webpage...'
    html = urllib.urlopen('http://disneyworld.disney.go.com/parks/magic-kingdom/calendar/')
    print 'Creating soup...'
    soup = BeautifulSoup(html)
    print 'Parsing webpage...'
    events = dict()
    strmonth = {'01': 'january', 
                '02': 'february', 
                '03': 'march', 
                '04': 'april', 
                '05': 'may', 
                '06': 'june', 
                '07': 'july', 
                '08': 'august', 
                '09': 'september', 
                '10': 'october', 
                '11': 'november', 
                '12': 'december'}

    #find the HTML for the month, based on the id (e.g. id=april2012)
    month_c = soup.find('div', attrs={'id':strmonth[str(desmonth)]+'2012'})
    #find all the day objects in the month
    parking_month = month_c.findAll('div', 'dayContainer')
    #regex to find the hours and hours types in parking_month
    hours = r'\d+:0{2}\s\w{2}\s-\s\d+:0{2}\s\w{2}'
    type = 'Park Hours|Extra Magic Hours'
    #iterate through all the day objects
    for item in parking_month:
        #pull out the date from the link (last 8 chars)
        date = item.find('a').get('href')[-8:]
        #if the month is outside of the desired range, ignore it and continue
        if date[4:6] != desmonth:
            continue
        #using regex, find all of the hours and hour types from hrs
        hrs = item.find('p', attrs={'class':'moreLink'}).text
        types = re.findall(type,str(hrs))
        times = re.findall(hours,str(hrs))
        #create a dict from the types and time (they should allign correctly)
        events = zip(types,times)
        #add the event dict into the dictionary for the park on that day
        parkh[str(formatdate(date))] = events
    print 'Done parsing website'
    
def exceledit(desmonth,readfile,outfile):
    #open excel sheet
    import datetime
    from xlutils.copy import copy
    print 'Opening excel sheet...'
    book = xlrd.open_workbook(readfile, on_demand=True, formatting_info=True)
    print 'Creating and editing new excel sheet...'
    wbook = copy(book)
    wbook.dates_1904 = book.datemode
    print 'Done creating new excel sheet'
    
    sh = book.sheet_by_index(0)
    timestyle = xlwt.easyxf('font: name Verdana; alignment: horizontal center, vertical bottom;')
    #iterate through dates in excel sheet
    for colnum in range(sh.ncols):
        if colnum in range(0,4):
            continue
        
        date = sh.cell_value(3, colnum)
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
                wbook.get_sheet(0).write(6, colnum, "")
                wbook.get_sheet(0).write(5, colnum, "")
                wbook.get_sheet(0).write(7, colnum, "")
                #set default start and close times. will likely be overwritten
                starttime = parsetime(format, '10:00 AM')
                closetime = parsetime(format, '10:00 PM')
                #iterate through hour segments for that day
                for x in parkh[format]:
                    #if regular hours, insert in "HOURS" row
                    if x[0] == 'Park Hours':
                        #set opening time user park hours
                        xtime = parsetime(format, x[1][0:8].rstrip())
                        if xtime < starttime:
                            starttime = xtime
                        
                        #set closing time using park hours
                        ytime = parsetime(format,x[1][-8:].rstrip())
                        """
                        If closing time is in the morning, it will come up as
                        before whenever the park opens, which will cause
                        problems. Therefore, if closing time (ytime) is before
                        opening time (e.g. 3:00 AM), then put closing time to
                        in the following day.
                        """
                        if ytime < starttime:
                            ytime += timedelta(days=1)

                        if ytime > closetime:
                            closetime = ytime

                        #write park hours to excel sheet
                        wbook.get_sheet(0).write(6, colnum, x[1].lower().replace(' ',''), timestyle)
                    #if extra magic hours, insert in respective row
                    if x[0] == 'Extra Magic Hours':
                        #insert in morning row
                        if int(x[1][0:1]) in range(2,9):
                            #set new opening time
                            xtime = parsetime(format,x[1][0:8].rstrip())
                            if xtime < starttime:
                                starttime = xtime
                            #write morning emh to excel sheet
                            wbook.get_sheet(0).write(5, colnum, x[1].lower().replace(' ',''),timestyle)
                        #insert in evening row
                        else:
                            ytime = parsetime(format,x[1][-8:].rstrip())
                            if ytime < starttime:
                                ytime += timedelta(days=1)
                            if ytime > closetime:
                                closetime = ytime
                            wbook.get_sheet(0).write(7, colnum, x[1].lower().replace(' ',''),timestyle)
                #edit wait times based on open/close times
                adjustwait(book,wbook,colnum,format,starttime,closetime)

    print 'Done editing. Now saving...'
    wbook.save(outfile)
    print outfile+' saved'

def adjustwait(book,wbook,colnum,day,starttime,closetime):
    #optional. just prints out respective open/close time for respective day
    print 'Opening time: ' + str(starttime)
    print 'Closing time: ' + str(closetime)
    """
    set cutoff time for current day/next day (i.e. times that are before 7:00AM
    are likely going to be early hour next day [e.g. 3:00 AM, etc])
    """
    cuttime = parsetime(day, '7:00 AM')
    
    sh = book.sheet_by_index(0)
    for rownum in range(sh.nrows):
        #skip the first four rows, as they contain no time information
        if rownum in range(0,4):
            continue
        #grab cell value (time)
        ctime = sh.cell_value(rownum, 3)
        if ctime:
            ntime = parsetime(day,ctime.replace('.',''))
            #if the time is before 7:00 AM, change to next day
            if ntime < cuttime:
                ntime += timedelta(days=1)
            #if the time is not within hours of operation, set the wait time
            if not starttime <= ntime < closetime:
                style = xlwt.easyxf('font: name Verdana; alignment: horizontal center; pattern: pattern solid, fore_color light-orange;')
                wbook.get_sheet(0).write(rownum,colnum,'CL',style)
    
if __name__ == '__main__':
    main()
