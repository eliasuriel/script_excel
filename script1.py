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

# Reemplaza los valores nulos en la columna 2 con ceros
df['Sep'] = df['Sep'].fillna(0)
columna2 = df['Sep']
columna1 = df['ResourceName']


Productivity1 = 0
Productivity2 = 0
horas = []
horas2 = []
EGB_Group = []
cont = 0
cont_empleados = 0

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
    elif 'EGB3' in columna_tareas or 'EGB8' in columna_tareas or 'EGB9' in columna_tareas or 'EGB10' in columna_tareas or 'EGB11' in columna_tareas or 'EGB12' in columna_tareas:
        cont = cont + 1
        cont_empleados = cont_empleados + 1
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

    if 'EGB3' in columna_tareas:
        EGB_Group.append('EGB3')
    elif 'EGB8' in columna_tareas:
        EGB_Group.append('EGB8')
    elif 'EGB9' in columna_tareas:
        EGB_Group.append('EGB9')
    elif 'EGB10' in columna_tareas:
        EGB_Group.append('EGB10')
    elif 'EGB11' in columna_tareas:
        EGB_Group.append('EGB11')
    elif 'EGB12' in columna_tareas:
        EGB_Group.append('EGB12')
    elif 'EGB' in columna_tareas:
        EGB_Group.append('EGB')
    #else:
     #   EGB_Group.append(' ')

#print(EGB_Group)  
print(cont_empleados)
#print("Horas: \n")
print(len(horas))
print(horas)
#print("OTRA \n" )
#print(horas2)
