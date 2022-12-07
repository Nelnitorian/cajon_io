# -*- coding: utf-8 -*-

import openpyxl
#from tabulate import tabulate


##############################################################
### Leer y grabar tablas con cabeceras en un rango (Los titulos de filas y columnas del rango se usan como claves en un diccionario anidado) Doble Diccionario
##############################################################

def Read_Excel_to_NesteDic(sheet, Range1, Range2): # Los datos en la hoja de cálculo deben haber sido formateados primero
    dict1={}
    multiple_cells = sheet[Range1:Range2]
    Aux={}
    Aux.update({0:'Empty'})

    #Primero vamos a leer la fila que contiene las cabeceras de las columnos, que seran claves del diccionario interno
    Column=0
    for cell in multiple_cells[0]:
        if Column >=1:
            Aux.update({Column:cell.value})
        Column=Column+1
    # Ahora pasamos a leer por filas desde la primera

    RowNumber=len(multiple_cells)
    for Row in range(1,RowNumber):
        dict2 = {}
        Column=0
        key = multiple_cells[Row][Column].value
        for cell in multiple_cells[Row]:
            if Column>=1:
                dict2.update({Aux[Column]:cell.value})
            Column=Column+1
        dict1.update({key:dict2})

    return dict1



def Read_Excel_to_NesteDic_tuple(sheet, Range1, Range2): # Los datos en la hoja de cálculo deben haber sido formateados primero
    dict1={}
    multiple_cells = sheet[Range1:Range2]
    Aux={}
    Aux.update({0:'Empty'})

    #Primero vamos a leer la fila que contiene las cabeceras de las columnos, que seran claves del diccionario interno
    Column=0
    for cell in multiple_cells[0]:
        if Column >=1:
            Aux.update({Column:cell.value})
        Column=Column+1
    # Ahora pasamos a leer por filas desde la primera

    RowNumber=len(multiple_cells)
    for Row in range(1,RowNumber):
        dict2 = {}
        Column=0
        key=tuple(int(x) for x in multiple_cells[Row][Column].value[1:-1].split(','))
        #key = tuple(map(int, elt[0].split(','))) for elt in multiple_cells[Row][Column].value
        for cell in multiple_cells[Row]:
            if Column>=1:
                dict2.update({Aux[Column]:cell.value})
            Column=Column+1
        dict1.update({key:dict2})

    return dict1
##############################################################

def Write_NesteDic_to_Excel(WB, name, sheet, Dict, Range1,Range2):

    multiple_cells = sheet[Range1:Range2]
    aux1=[]
    aux2=[]
    #Leyendo claves de filas
    aux1=getList(Dict)
    # Leyendo claves de columnas
    aux2=getList(Dict[aux1[0]])


    auxdic={}
    #Leyendo elementos del diccionario
    for i in Dict:
        for j in Dict[i]:
            auxdic.update({(i,j):Dict[i][j]})

    #Primero vamos a escribir la fila que contiene las cabeceras de las columnas, que seran claves del diccionario interno
    Column=0
    for cell in multiple_cells[0]:
        if Column==0:
            cell.value=' '
        if Column >=1:
            cell.value=aux2[Column-1]
        Column=Column+1

    # AHora escribimos las claves por columnas y el contenido
    RowNumber=len(multiple_cells)
    ColNumber=len(multiple_cells[0])
    for Row in range(1,RowNumber):
        multiple_cells[Row][0].value=aux1[Row-1]
        for j in range(1,ColNumber):
            a1=aux1[Row-1]
            a2=aux2[j-1]
            #print(auxdic[a1,a2])
            multiple_cells[Row][j].value=auxdic[a1,a2]
    WB.save(name)

def getList(dict):
    list = []
    for key in dict.keys():
        list.append(key)

    return list

##############################################################
# Leer y grabar listas en un rango

##############################################################
def Read_Excel_to_List(sheet,Range1, Range2):
    listaAux = []
    multiple_cells = sheet[Range1:Range2]
    for row in multiple_cells:
        for cell in row:
            listaAux.append(cell.value)

    return listaAux
##############################################################

def Write_List_to_Excel(wb, name, sheet, List1, Range1, Range2):
    multiple_cells = sheet[Range1:Range2]
    k=0
    for row in multiple_cells:
        for cell in row:
            cell.value=List1[k]
            k=k+1

    wb.save(name)

##################################################################
### Leer y grabar contenido de diccionarios sin keys en un rango

###################################################################

def Read_Excel_to_DicTable(sheet,Range1, Range2):
    Dict = []
    multiple_cells = sheet[Range1:Range2]
    i=1
    j=1
    for row in multiple_cells:
        for cell in row:
            Dict[i,j].update({(i,j):cell.value})
            j+=1
        i+=1

    return Dict

##############################################################
def Write_DicTable_to_Excel(wb, name, sheet, Dict, Range1, Range2):

    multiple_cells = sheet[Range1:Range2]
    Rows=len(multiple_cells)
    Columns=len(multiple_cells[0])
    aux=list(Dict.values())
    i=0
    for row in multiple_cells:
        for cell in row:
            cell.value=aux[i]
            i=i+1
            if i >= len(aux):
                break
        if i >= len(aux):
            break
    wb.save(name)
####################################################################

####################################################################



