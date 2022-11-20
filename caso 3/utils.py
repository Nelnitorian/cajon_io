# -*- coding: utf-8 -*-
import collections

import re
import pandas as pd
import json
import numpy as np

__author__ = "Juan José Rosendo, Eduardo Rubio, Belén María Lozano"
__credits__ = ["Dpto. Ing. de Organización, Universidad de Sevilla"]
"""
    Para crear estos scripts de utilidades se ha partido de forma vaga de los scripts propuestos en IOfunctionsExcel.py.
"""
def Read_Excel_to_List(sheet,Range1, Range2):
    listaAux = []
    multiple_cells = sheet[Range1:Range2]
    for row in multiple_cells:
        for cell in row:
            listaAux.append(cell.value)

    return listaAux

def read_dic_single_value_from_excel(sheet, ranges):

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

def parse_dic_values(dic, **kwargs):
    for key in dic.keys():
        dic[key] = parse_single_value(dic[key], **kwargs)
    return dic

def parse_single_value(string, char=','):
    frags = string.split(char)
    return tuple(json.loads(f"[\"{frags[0]}\", \"{frags[1]}\"]"))

def create_dictionary_vars(solver, primary_key, secondary_key):
    dic = create_empty_nested_dics(primary_key)
    for k1 in primary_key:
        for k2 in secondary_key:
            dic[k1][k2] = solver.NumVar(0, solver.infinity(), f"x_{k1}_{k2}")
    return dic

def create_empty_nested_dics(keys):
    lst = []
    for i in range(len(keys)):
        lst.append(dict())
    return dict(zip(keys, lst))

def create_dic(keys, values):
    return dict(zip(list(keys), list(values)))

def extract_list_value(dic):
    for key in dic.keys():
        dic[key] = dic[key][0]
    return dic

def write_nested_dicts_to_excel(WB, name, sheet, dic, range_list, header):
    if len(range_list)==1:
        range1, range2 = calculate_write_ranges_dic(dic, range_list[0])
    else:
        range1 = range_list[0]
        range2 = range_list[1]
    multiple_cells = sheet[range1:range2]

    primary_keys = list(dic.keys())
    secondary_keys = list(dic[list(dic.keys())[0]].keys())

    # Write data
    for i in range(len(primary_keys)):
        k1 = primary_keys[i]
        for j in range(len(secondary_keys)):
            k2 = secondary_keys[j]
            multiple_cells[i+1][j+1].value = dic[k1][k2]

    # Write headers
    multiple_cells[0][0].value = header
    for i in range(len(primary_keys)):
        multiple_cells[i+1][0].value = primary_keys[i]
    for j in range(len(secondary_keys)):
        multiple_cells[0][j+1].value = secondary_keys[j]

    WB.save(name)
    return


def write_list_to_excel(WB, name, sheet, lst, range_list, header):
    if len(range_list)==1:
        range1, range2 = calculate_write_ranges_lst(lst, range_list[0])
    else:
        range1 = range_list[0]
        range2 = range_list[1]
    multiple_cells = sheet[range1:range2]

    # Write data
    for i in range(len(lst)):
        multiple_cells[i+1][0].value = lst[i]

    # Write headers
    multiple_cells[0][0].value = header

    WB.save(name)
def calculate_write_ranges_dic(dic, start):
    l1 = len(list(dic.keys()))
    l2 = len(list(dic[list(dic.keys())[0]].keys()))
    r1 = start
    start_col = re.findall(r'\D+', r1)[0]
    start_row = re.findall(r'\d+', r1)[0]
    r2 = chr(ord(start_col)+l2)+str(int(start_row)+l1)
    return r1, r2

def calculate_write_ranges_lst(lst, start):
    l1 = len(list(lst))
    r1 = start
    start_col = re.findall(r'\D+', r1)[0]
    start_row = re.findall(r'\d+', r1)[0]
    r2 = start_col+str(int(start_row)+l1)
    return r1, r2
