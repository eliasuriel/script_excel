
import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import NamedStyle
from openpyxl.worksheet.table import Table, TableStyleInfo
 
import tkinter as tk
from tkinter import filedialog
import easygui


#easygui.msgbox("Choose your file to be proccessed (Excel File)")
#print("Asking for input file")
#root = tk.Tk()
#root.withdraw()

#file_path = filedialog.askopenfilename()
#file_path = "C:\\Users\\VEU1GA\\Documents\\Visual_Studio\\Script2"

# Reemplaza 'nombre_del_archivo.xlsx' con el nombre de tu archivo Excel
excel_file = pd.ExcelFile('PRIME.xlsx')
opx_file = pd.ExcelFile('opx_new_1.xlsx')
easygui.msgbox("The file was processed successfully")

anio =  int(easygui.enterbox("Type the year"))
str_anio = str(anio)

numhojas = 0

numhojas = int(easygui.enterbox("Input the number of the prime sheets (or months) do you want to work on"))
#numhojas = 1
hojas = []
hojas_2 = []
for i in range(numhojas):
   hojas.append(str(easygui.enterbox("In which prime file sheet (or month) do you want to work? sheet number:"  + str(i + 1))))

#PRIME#
#hojas.append("Associate_Jan")
#hojas.append("Associate_Feb")
#hojas.append("Associate_March")
#hojas.append("Associate_April")
#hojas.append("Associate_May")
#hojas.append("Associate_June")
#hojas.append("Associate_July")

#OPX
hojas_2.append("Table")
hojas_2.append("Sheet1")

#We are going to use only two names, because is the same name of the two columns for all the sheets
#nombrecolumna1 = str(easygui.enterbox("Name of the first column to work on? (coworkers,name of the jobs)"))
#nombrecolumna2 = str(easygui.enterbox("Name of the second column to work on? (hours)"))

#name of the columns of PRIME file
nombrecolumna1 ="Resource Name"
nombrecolumna2= "Hours"

Nombres_Tareas = []
hours = []
column_headers = []


cont = 0
cont_anio = 0
suma_prime = 0
anio_mes = ""
palabra = "BE-"
palabra2 = "PN-"
condicional = 0
suma_opx = 0
totalopx = 0
mesopx = ""




###########################################################
#                   Archivo Output
# #########################################################   
 
archivos = str(easygui.enterbox(msg="Write a name to the output file "))  
archivo  = archivos + ".xlsx"
 

easygui.msgbox("Enter the folder where you want the output file to be stored ")
directorio=str(easygui.diropenbox())


 
ruta_completa = os.path.join(directorio, archivo)
print(ruta_completa)
 
# Verificar si el archivo existe
 
if not os.path.exists(ruta_completa):
    workbook = openpyxl.Workbook()      #Crear un nuevo libro
    hoja = workbook.active              #Crear una nueva hoja
    hoja.title = str_anio               #Título de la hoja = anio
    workbook.save(ruta_completa)        #Guardar el libro
 
    easygui.msgbox(f"The file {archivo} was created in: {directorio}")
    
else:
    easygui.msgbox(f"The file {archivo} already exists in: {directorio}")
    workbook = openpyxl.load_workbook(ruta_completa)
    hoja = workbook.active              #Crear una nueva hoja
    hoja.title = str_anio               #Título de la hoja = anio
    workbook.save(ruta_completa)        #Guardar el libro
 
HOURS_PER_DAY =  float(easygui.enterbox("How many laboral hours does a day have? (8, 8.25...)"))
#HOURS_PER_DAY = 8.25

num_fila = 1
num_columna = 1



#Imprimir encabezados
#hoja.cell(row = num_fila)




#########################################################
#             Iteraciones y Calculos
#######################################################

for i in range(numhojas):
    df = excel_file.parse(hojas[i])
    df_2= opx_file.parse(hojas_2[0])

    
    easygui.msgbox("Answer the following box as follows, type 1 if you want to work in the month of January, 2 if in the month of February.....")
    mes = int(easygui.enterbox("Type the month in which you want to work in opx in number (must be the same month as prime)" + "This is the month of prime that you choose before:" + str(hojas[i])))    

    num_proyectos = int(easygui.enterbox("How many projects do you want to work on this month " + str(hojas[i]) + "?"))    

    #num_proyectos = 1 #num the     
    df = df.dropna(subset=[nombrecolumna1, nombrecolumna2])
    column_headers = df_2.columns.tolist() #   
    df_2[column_headers] = df_2[column_headers].fillna(0)

    if mes == 1:
        anio_mes = str_anio 
    elif mes == 2:
        anio_mes = str_anio + ".1"
    elif mes == 3:
        anio_mes = str_anio + ".2"
    elif mes == 4:
        anio_mes = str_anio + ".3"
    elif mes == 5:
        anio_mes = str_anio + ".4"
    elif mes == 6:
        anio_mes = str_anio + ".5"
    elif mes == 7:
        anio_mes = str_anio + ".6"
    elif mes == 8:
        anio_mes = str_anio + ".7"
    elif mes == 9:
        anio_mes = str_anio + ".8"
    elif mes == 10:
        anio_mes = str_anio + ".9"
    elif mes == 11:
        anio_mes = str_anio + ".10"
    elif mes == 12:
        anio_mes = str_anio + ".11"
    else:
        anio_mes = "0"


    for col_header in column_headers: #recorre los headers del opx-*89
        if str_anio in str(col_header) and anio_mes == str(col_header): #verifica si coincide 2023(o el anio que escriban) o si esta en algun header
            cont_anio += 1
            
            #print(cont_anio)#nomas era para ver cuanta veces entraba
            #print(col_header)#verificar que si fuera correcta la comparacion

            for j in range(num_proyectos):    
                proyecto = str(easygui.enterbox("Name of the project you want to work on? (Ford, Singer, Tesla, BRP, Zoox)"))
                num_areas = int(easygui.enterbox("For this project: " + str(proyecto) + "; how many areas would you like to analyze?"))
                #num_areas = 1
                
                for h in range (num_areas):
                    if suma_prime !=0:
                        suma_prime = 0
                    if suma_opx !=0:
                        suma_opx = 0
                    area = str(easygui.enterbox("Name of the area you want to work on this project (" + proyecto + ") (NET,CSW, ASW, etc)"))
                    
                    #print(df_2)
                    #print("Columnas de la hoja df_2:", column_headers)
                    #df_2[nom3] = df_2[nom3].fillna(0)
                    
                    for index, row in df.iterrows():
                            columna_tareas = row[nombrecolumna1]
                            columna_horas = row[nombrecolumna2]    

                            #print(columna_tareas)
                            #print (columna_horas)
                            cont += 1
                            if proyecto in str(columna_tareas) and area in str(columna_tareas):
                                print("entre")
                                suma_prime = suma_prime + columna_horas

                    for index, row in df_2.iterrows():                    
                        columna_1 = row[column_headers[0]]
                        columna_2 = row[col_header] 
                        if index == 0:
                            #easygui.msgbox(columna_2)  
                            mesopx = columna_2                     
                        
                        #print(columna_1)
                        #print(columna_2)

                        if str(columna_1).startswith(palabra or palabra2):
                            condicional = 0
                        
                        if proyecto in str(columna_1):
                            condicional = 1
                        
                        if condicional == 1 and area in str(columna_1):
                            suma_opx = suma_opx + columna_2*HOURS_PER_DAY     

                    easygui.msgbox("Mes Prime: " + str(hojas[i]) + "\nMes OPX: " + str(mesopx) + "\nProyecto: " + str(proyecto) + "\nArea: " + str(area)  + "\nHoras Prime: " + str(suma_prime) + "\nHojas OPX: " + str(suma_opx))
                    # easygui.msgbox(suma_prime)
                    # easygui.msgbox(suma_opx)        

                        






