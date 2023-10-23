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

df = df[df['Sep'] != 0]

cont_product = 0
cont_horas = 0
employees = []
horas = []
i = 2

WORK_DAYS = int(input("How many work days are in this report?: ") )

for index, row in df.iterrows():
    columna_tareas = row['ResourceName']
    columna_horas = row['Sep']
    print(columna_tareas)
    print(columna_horas)
    if columna_horas == 0:
        print()
    if columna_tareas.startswith("   P_"):
        #print("entre")
        cont_product = cont_product + columna_horas
    elif columna_tareas.endswith('-MS)') or columna_tareas.endswith('-MX)') or columna_tareas.endswith('-SX)'):
        if cont_product!=0:
            cont_product = cont_product/148
            horas.append(cont_product)
            cont_product = 0
    

    



#print(cont)

print(horas)
 
#print(df)