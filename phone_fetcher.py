import urllib.request, copy

def get_names_and_numbers(html):
    nameNumList = [[],[]]
    
    nameStart = "ContactName"
    nameEnd = "</a></span>"

    numStart = "ContactPhone"
    
    appending = False

    for i in range(len(html)):
        tmpstr = ""

        if html[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html[i:i + len(nameEnd)] == nameEnd:
            j = i - 1

            while html[j] != ">":
                tmpstr += html[j]
                j -= 1

            nameNumList[0].append(tmpstr[::-1])
            appending = False
    
    appending = False

    for i in range(len(html)):
        if html[i:i + len(numStart)] == numStart:
            appending = True
        
        if appending and html[i - 1] == ">":
            nameNumList[1].append(html[i:i + 14])
            appending = False
    
    return nameNumList

def read_site_to_file(url):
    try:
        site = urllib.request.urlopen(url)
    except urllib.error.URLError:
        return
    
    file = site.read().decode("utf-8")
    site.close()
    
    return file

if __name__ == '__main__':
    pcode = input("Postal Code (no spaces): ")
    url = "http://www.canada411.ca/search/?stype=pc&pc=" + pcode
    html = read_site_to_file(url)
    contacts = get_names_and_numbers(html)
    
    for i in range(len(contacts[0])):
        print(contacts[0][i] + ": " + contacts[1][i])