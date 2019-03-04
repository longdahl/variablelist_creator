

from requests import get
from bs4 import BeautifulSoup
import xlsxwriter
import pandas as pd
import csv
import urllib

url = 'https://www.dst.dk/da/TilSalg/Forskningsservice/Data/Register_Variabeloversigter'

response = get(url)




soup = BeautifulSoup(response.content, "lxml")

#path = "C:/Users/mikkel-bj/Desktop/datamanager/script/from_dst/"
path = "D:/datamanager/script/from_dst/"
def crawler(path):

    for a in soup.find_all('a', href=True):

        if a['href'].__contains__("extranet"):
            if a['href'].__contains__("http://"):
                continue
            start_index = a['href'].find("Variabellister/")
            length = len("Variabellister/")
            total = start_index + length
            name = a['href']
            name = name[total::]
            name = name.split(' -')[0]
            #name = a['href'][total:total+4]
            #name = name.replace(" ","")
            #name = name.replace("-", "")
            #name = name.replace("_", "")
            print(name)
            print(len(name))
            url2 = 'https://www.dst.dk' + a['href']
            url2 = url2.replace(" ","%20")
            url2 = url2.replace("Æ","%C3%86")
            url2 = url2.replace("æ","%C3%A6")
            url2 = url2.replace("ø","%C3%B8")
            url2 = url2.replace("å","%C3%A5")
            url2 = url2.replace("–","%E2%80%93") #note this charachter is NOT a normal "-"!!
            url2 = url2.replace("Ø","%C3%98")
            url2 = url2.replace("§", "%C2%A7")

            print(url2)
            try:
                dfs = pd.read_html(url2)
                dfs = dfs[0]
                dfs.to_excel(path + name + "2.xlsx")
            except (urllib.error.HTTPError, IndexError):
                print("error")


crawler(path)


"""
        response2 = get(url2)
        soup2 = BeautifulSoup(response2.content, "lxml")
        tables = soup2.find_all('table')
        table = tables[1]
        pd.read_html(table)
        headers = [th.text.encode("utf-8") for th in table.select("tr th")]
        with open("C:/Users/mikkel-bj/Desktop/datamanager/script/from_dst/" + name + ".csv", "w") as f:
            wr = csv.writer(f)
            wr.writerow(headers)
            wr.writerows([[td.text.encode("utf-8") for td in row.find_all("td")] for row in table.select("tr + tr")])
            f.flush()

        #response2 = get(url2)
        #soup2 = BeautifulSoup(response2.content, "lxml")
        #register_containers3 = soup2.find_all('div')
        #print(register_containers3)
        #workbook.close()

"""



"""
html_soup = BeautifulSoup(response.text, 'html.parser')
type(html_soup)
soup = BeautifulSoup(response.content, "lxml")


register_containers = html_soup.find_all(href=True)
register_containers2 = soup.find_all('div')
print(type(register_containers))
print(register_containers)

"""

#print(type(register_containers2))
#print(register_containers2)

#get div
#get href
"""


from bs4 import BeautifulSoup,SoupStrainer
import httplib2

http = httplib2.Http()

status,response = http.request('https://www.dst.dk/da/TilSalg/Forskningsservice/Data/Register_Variabeloversigter')


for link in BeautifulSoup(response, parse_only=SoupStrainer('a')):
    if link.has_attr('href'):
        print(link['href'])

"""









