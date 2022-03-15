# -*- coding: utf-8 -*-
"""
Created on Fri Jul  3 20:36:44 2020
"https://https://www.siogranos.com.ar/Consulta_publica/operaciones_informadas_exportar.aspx")
@author: feder
"""

#Selenium


##################
# Descargo archivo desde SIOGRANOS usando Selenium
#import sys
import time
#from datetime import datetime, timedelta
import datetime
from selenium import webdriver
#from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import os 

os.chdir("C:/Users/fdiyenno/Mi unidad/10. Python/Back up scripts spyder/3.Siogranos_Selenium")

try:
    PATH = "C:/Users/fdiyenno/Mi unidad/10. Python/Back up scripts spyder/3.Siogranos_Selenium/chromedriver.exe"
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920x1080") #I added this
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory": r"C:\Users\fdiyenno\Mi unidad\10. Python\Back up scripts spyder\3.Siogranos_Selenium\Download"}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path=PATH,options=chrome_options)

#Si hay error me descargo el nuevo driver del chrome
except Exception as err:
    Version = err
    
    #Imprimo el error y lo almaceno en una variable
    import sys
    import io
    old_stdout = sys.stdout
    new_stdout = io.StringIO()
    sys.stdout = new_stdout
    print(Version)
    output = new_stdout.getvalue()
    sys.stdout = old_stdout
    print(output)
    
    #Saco la versión del error
    import re
    my_string=output
    try:
        B = my_string.split("Current browser version is ",1)[1] 
    except:
        ""
    version = re.search(r'\d+', B).group()
    
    import requests
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_'
    url_file = 'https://chromedriver.storage.googleapis.com/'
    file_name = 'chromedriver_win32.zip'
    
    version = version
    version_response = requests.get(url + version)
    
    if version_response.text:
        file = requests.get(url_file + version_response.text + '/' + file_name)
        with open(file_name, "wb") as code:
            code.write(file.content)
    
    import zipfile
    with zipfile.ZipFile(file_name,"r") as zip_ref:
        zip_ref.extractall("")

    PATH = "C:/Users/fdiyenno/Mi unidad/10. Python/Back up scripts spyder/3.Siogranos_Selenium/chromedriver.exe"
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920x1080") #I added this
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory": r"C:\Users\fdiyenno\Mi unidad\10. Python\Back up scripts spyder\3.Siogranos_Selenium\Download"}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(executable_path=PATH,options=chrome_options)

#driver = webdriver.Chrome(PATH)

dt = datetime.datetime.today()
print (dt.day)

# Dependiendo de la hora seleccionamos el dìa actual o el anterior
#ahora = datetime.datetime.now().hour
#if ahora > 10 and ahora < 20:
dia = str(dt.day)
#else:
#    dia = str(dt.day-1)


driver.get("https://www.siogranos.com.ar/Consulta_publica/operaciones_informadas_exportar.aspx")

print(driver.title)

#search = driver.find_element_by_name("txtFechaConcertacionDesde")
search = driver.find_element_by_name("txtFechaOperacionDesde")
search.click()
time.sleep(1)
fecha = driver.find_element_by_link_text("1")
fecha.click()
#search = driver.find_element_by_name("txtFechaConcertacionHasta") Necesitamos saber no fecha de concertqacion sino fecha de operación
search = driver.find_element_by_name("txtFechaOperacionHasta")
search.click()
time.sleep(1)
fecha = driver.find_element_by_link_text(dia)
fecha.click()
time.sleep(1)
archivo = driver.find_element_by_name("btn_generar_csv")
archivo.click()

 
#time.sleep(15)
#driver.quit()
paths = r"C:\Users\fdiyenno\Mi unidad\10. Python\Back up scripts spyder\3.Siogranos_Selenium\Download\operaciones_informadas.csv"

import os.path
import time

while not os.path.exists(paths):
    time.sleep(10)

if os.path.isfile(paths):
    driver.quit()
    print("i")# read file
else:
   raise ValueError("%s isn't a file!" % paths)

# Python program to convert a list to string 
    
# Function to convert   
def listToString(s):  
    
    # initialize an empty string 
    str1 = ""  
    
    # traverse in the string   
    for ele in s:  
        str1 += ele   
    
    # return string   
    return str1  
        
# Driver code     

print(listToString(paths))
operaciones = listToString(paths)

##################
# Modificamos el excel y lo guardamos en un lugar
import pandas as pd

data = pd.read_csv(operaciones, sep = ";", encoding = 'utf-16-le')

# dropping null value columns to avoid errors 
data.dropna(inplace = True, how = "all")
data = data.iloc[:, :-1]

# new data frame with split value columns 
new = data["FECHA OPERACION"].str.split(" ", n = 1, expand = True) 
# making separate first name column from new data frame 
data["FECHA OPERACION"]= new[0]
# making separate last name column from new data frame 
data["hora"]= new[1] 
new = data["FECHA CONCERTACION"].str.split(" ", n = 1, expand = True)
data["FECHA CONCERTACION"]= new[0]
new = data["FECHA ENTR. DESDE"].str.split(" ", n = 1, expand = True)
data["FECHA ENTR. DESDE"]= new[0]
new = data["FECHA ENTR. HASTA"].str.split(" ", n = 1, expand = True) 
data["FECHA ENTR. HASTA"]= new[0]

# Dropping old Name columns 
#data.drop(columns =["FECHA OPERACION"], inplace = True) 
  
# df display 
data 

cols = data.columns.tolist()
cols.insert(1, cols.pop(cols.index('hora')))
data = data.reindex(columns= cols)

data.to_csv('operaciones_informadas.csv', index=False, encoding='utf-16-le',sep=';')
data.to_excel('operaciones_informadas.xlsx', index=False, sheet_name='operaciones_informadas')


##############
#Subo el archivo al sharepoint

from shareplum import Site
from shareplum import Office365
#from config import config
from shareplum.site import Version

authcookie = Office365('https://bcr.sharepoint.com', username='pbi.diyee.dev@bcr.com.ar', password='BCR.PowerBi.2001').GetCookies()
site = Site('https://bcr.sharepoint.com/sites/DIyEE587/', version=Version.v2019, authcookie=authcookie, timeout = 240)
folder = site.Folder('Documentos%20compartidos/General/DIYEE%20-%20IYEE%20(Privado)/SIO-granos/Base_Siogranos')
#folder.upload_file(file_name, 'operaciones_informadas.csv')

import calendar
#from datetime import datetime
import datetime
d = datetime.date.today()

ahora = datetime.datetime.now()
ahora_1 = ahora.strftime('%m/%d/%Y, %H:%M:%S')
mes = datetime.datetime.now().month
año = datetime.datetime.now().year
_, num_days = calendar.monthrange(año,mes)
mes = d.strftime('%m')

tipo_archivo = "xlsx"
"Base SIO granos 2015_06_01 al 2015_06_30"
archivo_sio = "Base SIO granos %s_%s_01 al %s_%s_%s.%s" % (año,mes,año,mes,num_days,tipo_archivo) 


if tipo_archivo == "xlsx":
    import io
    
    file = io.open('operaciones_informadas.xlsx', "rb", buffering = 0)
    folder.upload_file(file, archivo_sio)
else:
    #import csv
    with open('operaciones_informadas.csv', 'rb') as csvf:
            fileContent = csvf.read()
    folder.upload_file(fileContent, "filename.csv")

# Creamos un log de intentos realizados
carga_completa = "Carga exitosa %s" % (ahora_1) 

with open("operaciones_informadas.txt", "a+") as file_object:
    # Append 'hello' at the end of file
    file_object.write("\n")
    file_object.write("%s \ %s" % (carga_completa,archivo_sio))
print(carga_completa)

import os
os.remove(paths)



