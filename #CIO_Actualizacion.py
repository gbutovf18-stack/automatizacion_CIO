#CIO_Actualizacion 

import pandas as pd
import numpy as np
import win32com.client as wc
from datetime import date
from dateutil.parser import parse
import logging 
import time 
from selenium import webdriver

# url parameters
path = 'D:/Users/Gabriel/Documents/Prueba_CIO/'
file_path = "Adelantamiento.xlsx"
path_validator_base= "C:/Users/Gabriel/ownCloud/Bases/Consolidado_TOA.xlsx"
path_driver= "C:/Program Files (x86)/chromedriver.exe"
driver = webdriver.Chrome(path_driver)

#Aeguramos que el consolidado está actualizado
print ("Verificando que el archivo está actualizado...")
FILE_1 =  pd.read_excel(path_validator_base)


#Verificamos la min_fecha del archivo
print("Verificando la min_fecha del archivo...")
min_date= parse(FILE_1['Fecha'].min())
print('La min_fecha del archivo es: ' + min_date.strftime("%Y-%m-%d"))


if(min_date.strftime("%Y-%m-%d")==date.today().strftime("%Y-%m-%d")):
    print("El archivo está actualizado...")
    FILE =  wc.Dispatch("Excel.application")
    FILE.visible = True
    print("Abriendo archivo Adelantamiento_CIO..")
    workbook = FILE.Workbooks.open(path+file_path)
    print("Actualizando archivo Adelantamiento_CIO..")
    workbook.RefreshAll()
    FILE.CalculateUntilAsyncQueriesDone()
    print("El archivo fue actualizado...")
    print("Guardando archivo...")
    workbook.Save()
    
    

else:
    print("El archivo no está actualizado...")
    logging.warning("Pendiente por validacion.")
