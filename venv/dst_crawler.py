

from requests import get
from bs4 import BeautifulSoup
import xlsxwriter
import pandas as pd
import csv
import urllib

url = 'https://www.dst.dk/da/TilSalg/Forskningsservice/Data/Register_Variabeloversigter'
response = get(url)
soup = BeautifulSoup(response.content, "lxml")

arbejde = 1

if arbejde == 1:
    path = "C:/Users/mikkel-bj/Desktop/datamanager/script/from_dst/"
else:
    path = "C:\\Users\\Mikkel\Desktop\\arbejde\\Project database\\datamanager\\script\\from_dst\\"
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

            #code to update a specific register:
            #if name != "AKAS":
            #    print("not this" + name)
            #    continue


            #name = a['href'][total:total+4]
            #name = name.replace(" ","")
            #name = name.replace("-", "")
            #name = name.replace("_", "")
            print(name)
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
            except (urllib.error.HTTPError, IndexError,ValueError):
                print("error")


crawler(path)







