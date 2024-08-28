from openpyxl import*
from gurobipy import*

# Carga de Parametros
book = load_workbook("Punto1Datos.xlsx")

beneficiosSheet=book["Beneficios"]
comunicacionesSheet=book["Comunicaciones"]
costosSheet=book["Costos"]


#Definicion de Parametros
b = {} #Beneficios
for fila in range(2,4):
    for columna in range (2,6):
        sede = beneficiosSheet.cell(1,columna).value
        print(sede)
        
 

