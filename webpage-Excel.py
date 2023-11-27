# Extract all the info from the website and place into a spreadsheet.
# The data includes fields like Name, Active since, Industry, Investing status, and Contact information for each of the companies.
# Split the contact info into multiple columns: Contact person(name), Address, city, state, zip, phone number, email.

import urllib.request,urllib.error, urllib.parse
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook

def createTable(tags):

#Extracting text from tags

    tagsList = list()
    for tag in tags:
        childList = list()      
        for child in tag.children:
            if len(child.text.strip())!=0:
                childList.append(child.text.rstrip())
        tagsList.append(childList)

# Split the contact info into multiple columns: contact name, address, city, state, zip, email, phone number

    for i in tagsList:
        contactInformation = i[4].split('\n')
        
        if len(contactInformation)>1:
            cityStateZip = contactInformation[2]    #'Chicago, IL 60606'
            zip = cityStateZip[-5:]
            state = cityStateZip[-8:-6]
            city = cityStateZip[:-10]
            contactInformation.pop(2)
            contactInformation.insert(2,city)
            contactInformation.insert(3,state)
            contactInformation.insert(4,zip)
        
        i.pop(4)
        for c in contactInformation:
            i.append(c)
     
    tagsList.pop(0)                 #remove table header
    return tagsList


# serviceurl = input('Enter url:')
serviceurl = 'https://www.sba.gov/funding-programs/investment-capital/sbic-directory?industry=All&status=All&page='
page = '0'

wb = Workbook()                     #create an Excel workbook and spreadsheet
ws = wb.active
header = ['Name','Active since','Industry','Investing status','Contact person','Address','City','State','Zip','Email','Phone']       #create table header
ws.append(header)

while True:
    url = serviceurl+page
    html = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(html,'html.parser')
    tags = soup('tr')
    if len(tags)<1:                         #check if there is a table on the page, if no, break
        break
    print('Page',page)
    
    tagsList = createTable(tags)           #call function which prepares data to be placed into a spreadsheet

    for row in tagsList:                    #Place data into a spreadsheet.
        ws.append(row)

    time.sleep(1)
    page = str(int(page)+1)                 #go to the next Page

wb.save('Data.xlsx')
