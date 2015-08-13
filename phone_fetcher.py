import urllib.request, sys
from xlwings import *

def transpose(matrix):
    return [[matrix[j][i] for j in range(len(matrix))] for i in range(len(matrix[0]))]

def get_details(html411):
    detList = [[],[],[]]
    
    nameStart = "ContactName"
    nameSubStart = 'title="'
    
    addStart = "ContactAddress"

    numStart = "ContactPhone"
    
    appending = False

    for i in range(len(html411)):
        if html411[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html411[i:i + len(nameSubStart)] == nameSubStart:
            tmpstr = ""
            j = i + len(nameSubStart)

            while html411[j:j + len('">')] != '">':
                tmpstr += html411[j]
                j += 1

            detList[0].append(tmpstr)
            appending = False
    
    appending = False

    for i in range(len(html411)):
        if html411[i:i + len(addStart)] == addStart:
            appending = True
        
        if appending and html411[i] == ">":
            tmpstr = ""
            j = i + 1

            while html411[j:j + len('</span>')] != '</span>':
                tmpstr += html411[j]
                j += 1

            detList[1].append(tmpstr)
            appending = False
    
    appending = False

    for i in range(len(html411)):
        if html411[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html411[i - 1] == ">":
            detList[2].append(html411[i:i + 14])
            appending = False
    
    return transpose(detList)

def get_riding(htmlParl):
    ridingStart = 'ctl00_cphContent_repMP_ctl00_lblYellowBar'
    
    appending = False
    start = 0

    for i in range(len(htmlParl)):
        if htmlParl[i:i + len(ridingStart)] == ridingStart:
            i += len(ridingStart)
            appending = True
        
        if appending and htmlParl[i] == ",":
            return htmlParl[start:i]
        
        if appending and htmlParl[i] == ">":
            tmpstr = ""
            j = i + 1

            while htmlParl[j] != ',':
                tmpstr += htmlParl[j]
                j += 1

            return tmpstr.replace("--", "-")

    return ""

def read_cities():
    cities = []
    
    with open('cities.txt') as f: textlist = f.read().splitlines()
    
    for city in textlist:
        city = city.split(",")

        cities.append(city[0].split(":")[1][1:])
    
    return cities

def read_site_to_file(url):
    try:
        site = urllib.request.urlopen(url)
    except urllib.error.URLError:
        return
    
    file = site.read().decode().replace("&#039;", "'")
    site.close()
    
    return file

def excel_style(row, col):
    result = []
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    while col:
        col, rem = divmod(col - 1, 26)
        result[:0] = alpha[rem]

    return ''.join(result) + str(row)

def print_wb(contacts, pcode, row):
    Range(excel_style(row + 1, 1)).value = contacts[0].split(" ")[0]
    Range(excel_style(row + 1, 2)).value = ' '.join(contacts[0].split(" ")[1:])

    Range(excel_style(row + 1, 3)).value = contacts[2].replace("(", "").replace(")", "").replace(" ", "-")
    
    cities = read_cities()
    matches = [[],[]]

    for i in range(len(contacts[1])):
        for j in range(len(contacts[1])):
            if contacts[1][i:j] in cities:
                matches[0].append(contacts[1][:i - 1])
                matches[1].append(contacts[1][i:j])
    
    matches = transpose(matches)
    
    match = ["",""]

    for tmpmatch in matches:
        if len(tmpmatch[1]) > len(match[1]):
            match = tmpmatch

    Range(excel_style(row + 1, 4)).value = match[0]
    Range(excel_style(row + 1, 5)).value = match[1]

    Range(excel_style(row + 1, 6)).value = pcode
    Range(excel_style(row + 1, 7)).value = contacts[3]

if __name__ == '__main__':
    pcode = input("Postal Codes: ").split(",")
    
    domain411 = "http://www.canada411.ca/search/?stype=pc&pc="
    domainParl = "http://www.parl.gc.ca/Parlinfo/Compilations/HouseOfCommons/MemberByPostalCode.aspx?Menu=HOC&PostalCode="
    
    contacts = []

    for i in range(len(pcode)):
        tmpcode = pcode[i].split(" ")
        html411 = read_site_to_file(domain411 + tmpcode[0] + "+" + tmpcode[1])
        
        try:
            contacts.append(get_details(html411))
        except TypeError:
            break
    
    for i in range(len(pcode)):
        tmpcode = pcode[i].split(" ")
        htmlParl = read_site_to_file(domainParl + tmpcode[0] + tmpcode[1])     

        for j in range(len(contacts[i])):
            contacts[i][j].append(get_riding(htmlParl))

    wb = Workbook()
    
    k = 0

    for i in range(len(contacts)):
        for j in range(len(contacts[i])):
            print_wb(contacts[i][j], pcode[i], k)
            
            k += 1
