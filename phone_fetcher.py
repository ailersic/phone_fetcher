import urllib.request
from xlwings import *

def get_details(html):
    detList = [[],[],[]]
    
    nameStart = "ContactName"
    nameSubStart = 'title="'
    
    addStart = "ContactAddress"

    numStart = "ContactPhone"
    
    appending = False

    for i in range(len(html)):
        if html[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html[i:i + len(nameSubStart)] == nameSubStart:
            tmpstr = ""
            j = i + len(nameSubStart)

            while html[j:j + len('">')] != '">':
                tmpstr += html[j]
                j += 1

            detList[0].append(tmpstr)
            appending = False
    
    appending = False

    for i in range(len(html)):
        if html[i:i + len(addStart)] == addStart:
            appending = True
        
        if appending and html[i] == ">":
            tmpstr = ""
            j = i + 1

            while html[j:j + len('</span>')] != '</span>':
                tmpstr += html[j]
                j += 1

            detList[1].append(tmpstr)
            appending = False
    
    appending = False

    for i in range(len(html)):
        if html[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html[i - 1] == ">":
            detList[2].append(html[i:i + 14])
            appending = False    
    
    return detList

def read_site_to_file(url):
    try:
        site = urllib.request.urlopen(url)
    except urllib.error.URLError:
        return
    
    file = site.read().decode("utf-8")
    site.close()
    
    return file

def excel_style(row, col):
    result = []
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]

    return ''.join(result) + str(row)

if __name__ == '__main__':
    pcode = input("Postal Codes (no spaces, separated by commas): ").split(",")
    
    domain = "http://www.canada411.ca/search/?stype=pc&pc="
    
    contacts = []

    for i in range(len(pcode)):
        html = read_site_to_file(domain + pcode[i])
        contacts.append(get_details(html))
    
    wb = Workbook()
    for i in range(len(contacts)):
        Range(excel_style(1, 3 * i + 1)).value = pcode[i]
        
        for j in range(len(contacts[i][0])):
            Range(excel_style(j + 2, 3 * i + 1)).value = contacts[i][0][j]
            Range(excel_style(j + 2, 3 * i + 2)).value = contacts[i][1][j]
            Range(excel_style(j + 2, 3 * i + 3)).value = contacts[i][2][j]
