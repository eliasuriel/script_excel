import pandas as pd
import numpy as np
import math
import os
import csv

import tkinter as tk
from tkinter import filedialog

#print("Asking for input file")
#root = tk.Tk()
#root.withdraw()

#file_path = filedialog.askopenfilename()

#print("Reading File")
#df = pd.read_excel("Script_Input.xlsx") 

# Reemplaza 'nombre_del_archivo.xlsx' con el nombre de tu archivo Excel
excel_file = pd.ExcelFile('Script_Input.xlsx')

# Reemplaza 'nombre_de_la_hoja' con el nombre de la hoja en tu archivo Excel
df = excel_file.parse('Sheet1')

# Columna que contiene la información que buscas
columna_tareas = 'ResourceName'


# Accede a los datos de cada columna

# Reemplaza los valores nulos en la columna 2 con ceros
df['Sep'] = df['Sep'].fillna(0)
columna2 = df['Sep']
columna1 = df['ResourceName']

Productivity1 = 0
Productivity2 = 0
horas = []
horas2 = []
cont = 0

WORK_DAYS = int(input("How many work days are in this report?: ") )

for index, row in df.iterrows():
    columna_tareas = row['ResourceName']
    columna_horas = row['Sep']
    #print(columna_tareas)
    #print(columna_horas)
    if columna_tareas.startswith("   P_"):
        #print("entre")
        Productivity1 = Productivity1 + columna_horas
        Productivity2 = Productivity2 + columna_horas
    elif columna_tareas.endswith('-MS)') or columna_tareas.endswith('-MX)') or columna_tareas.endswith('-SX)'):
        cont = cont + 1
        if Productivity1 !=0:
            Productivity1 = Productivity1/148
            horas.append(Productivity1)
            Productivity1 = 0
        else:
            if(cont >= 2):
                horas.append(Productivity1)
        if  Productivity2 !=0:
            Productivity2 = Productivity2/148
            horas2.append(Productivity2)
            Productivity2= 0
    
#df = df[df['Sep'] != 0]


        

#print(cont)

print(horas)


print("OTRA")
 
print(horas2)
#print(df)