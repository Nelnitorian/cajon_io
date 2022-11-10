# -*- coding: utf-8 -*-
import collections

import openpyxl
import re
import pandas as pd
import json
import numpy as np
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



def Read_Dic_List_Value_from_Excel(sheet, ranges):
    """
    La idea detrás de esta función es usar la primera columna como claves de un diccionario y añadir los valores
    de la fila como una lista.
    :param sheet:
    :param ranges:
    :return:
    """
    pass
    # Get keys intervals:
    # range1 = ranges[0]
    # range2 = ranges[1]
    # key_r1 = re.findall(r'\D+', range1)[0]+re.findall(r'\d+', range1)[0]
    # key_r2 = re.findall(r'\D+', range1)[0]+re.findall(r'\d+', range2)[0]
    # val_r1 = re.findall(r'\D+', range2)[0]+re.findall(r'\d+', range1)[0]
    # val_r2 = re.findall(r'\D+', range2)[0]+re.findall(r'\d+', range2)[0]
    # keys = Read_Excel_to_List(sheet, key_r1, key_r2)
    # vals = Read_Excel_to_List(sheet, val_r1, val_r2)
    # dic = dict(zip(keys, vals))
    # multiple_cells_keys = sheet[range1_key:range2_key]
    # # # Init dictionary's keys
    # for cell in multiple_cells_keys:
    #     dictionary[cell[0].value] = []
    # diminished_range1 = chr(ord(re.findall(r'\D+', range1)[0])+1)+re.findall(r'\d+', range1)[0]
    #
    # multiple_cells_keys = sheet[diminished_range1:range2]
    # for row in multiple_cells_keys:
    #     for cell in row:
    #         dictionary[getKeyFromRange(cell.coordinate, range1)].append(cell.value)
    # return dic

def Read_Dic_Single_Value_from_Excel(sheet, ranges):
    """
    La idea detrás de esta función es usar la primera columna como claves de un diccionario y añadir los valores
    de la fila como un único elemento.
    :param sheet:
    :param ranges:
    :return:
    """

    # dictionary = Read_Dic_List_Value_from_Excel(sheet, ranges)
    # for key in dictionary.keys():
    #     dictionary[key] = dictionary[key][0]
    range1 = ranges[0]
    range2 = ranges[1]
    key_r1 = re.findall(r'\D+', range1)[0]+re.findall(r'\d+', range1)[0]
    key_r2 = re.findall(r'\D+', range1)[0]+re.findall(r'\d+', range2)[0]
    val_r1 = re.findall(r'\D+', range2)[0]+re.findall(r'\d+', range1)[0]
    val_r2 = re.findall(r'\D+', range2)[0]+re.findall(r'\d+', range2)[0]
    keys = Read_Excel_to_List(sheet, key_r1, key_r2)
    vals = Read_Excel_to_List(sheet, val_r1, val_r2)
    dic = dict(zip(keys, vals))
    return dic


def getKeyFromRange(coord, r1):
    return re.findall(r'\D+', r1)[0]+re.findall(r'\d+', coord)[0]

def read_from_excel_to_dataframe(excel_name, sheet_name, range):
    range_param = re.findall(r'\D+', range[0])[0]+":"+re.findall(r'\D+', range[1])[0]
    skiprows_param = int(re.findall(r'\d+', range[0])[0])-1
    nrows_param = int(re.findall(r'\d+', range[1])[0]) - skiprows_param
    return fill_df_values(pd.read_excel(excel_name, sheet_name, skiprows = skiprows_param, nrows=nrows_param,  usecols= range_param).dropna(how='all'))

def fill_df_values(df):
    for key in df.keys():
        aux_lst = fill_list_values(list(df[key]))
        if not collections.Counter(list(df[key])) == collections.Counter(aux_lst):
            df[key] = aux_lst
    return df

def fill_list_values(lst):
    for i in range(len(lst)-1):
        if lst[i+1] is np.nan:
            lst[i+1] = lst[i]
    return lst

def parse_dic_values(dic, char):
    for key in dic.keys():
        dic[key] = parse_single_value(dic[key])
    return dic

def parse_single_value(string):
    frags = string.split(',')
    return tuple(json.loads(f"[\"{frags[0]}\", \"{frags[1]}\"]"))

def create_dictionary_vars(solver, primary_key, secondary_key):
    dic = dict(zip(primary_key, [{}]*len(primary_key)))
    for k1 in primary_key:
        for k2 in secondary_key:
            dic[k1][k2] = solver.NumVar(0, solver.infinity(), f"x_{k1}_{k2}")
    return dic

def create_dic(keys, values):
    return dict(zip(list(keys), list(values)))

def extract_list_value(dic):
    for key in dic.keys():
        dic[key] = dic[key][0]
    return dic

# def create_dic_of_dic(key1, key2, val):
#     dic= dict(zip(key1, [{}]*len(key1)))
#     for k1 in dic.keys():
#         for k2 in key2:
#             dic[k1][k2] = val.