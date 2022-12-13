# -*- coding: utf-8 -*-
import collections

import re
import pandas as pd
import json
import numpy as np
from math import sqrt
from datetime import datetime

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

    key_r1 = extract_letters(range1)+extract_numbers(range1)
    key_r2 = extract_letters(range1)+extract_numbers(range2)
    val_r1 = extract_letters(range2)+extract_numbers(range1)
    val_r2 = extract_letters(range2)+extract_numbers(range2)

    keys = Read_Excel_to_List(sheet, key_r1, key_r2)
    vals = Read_Excel_to_List(sheet, val_r1, val_r2)
    dic = dict(zip(keys, vals))
    return dic

def read_from_excel_to_dataframe(excel_name, sheet_name, range):
    range_param = extract_letters(range[0])+":"+extract_letters(range[1])
    skiprows_param = int(extract_numbers(range[0]))-1
    nrows_param = int(extract_numbers(range[1])) - skiprows_param
    return fill_df_values(pd.read_excel(excel_name, sheet_name, skiprows = skiprows_param, nrows=nrows_param,  usecols= range_param).dropna(how='all'))

def fill_df_values(df):
    for key in df.keys():
        aux_lst = fill_list_values(list(df[key]))
        if not collections.Counter(list(df[key])) == collections.Counter(aux_lst):
            df[key] = aux_lst
    return df.fillna(0)

def fill_list_values(lst):
    for i in range(len(lst)-1):
        if lst[i+1] is np.nan:
            lst[i+1] = lst[i]
    return lst

def parse_dic_values(dic, **kwargs):
    for key in dic.keys():
        dic[key] = parse_single_value(dic[key], **kwargs)
    return dic

def extract_indexes_with_value(df):
    matrix = []
    for i in range(len(df)):
        a = df.loc[i, :]
        b = a[a == 1]
        matrix.append(list(b.keys()))
    return matrix

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

def create_list_empty_nested_dics(length):
    lst = []
    for i in range(length):
        lst.append({})
    return lst

def calculate_list_date_difference(unparsed_date_lst):
    aux_lst = []
    now = datetime.now()
    for then in unparsed_date_lst:
        delta = now-parse_date(then)
        aux_lst.append(delta.days)
    return aux_lst

def parse_date(date):
    return date.to_pydatetime()


def extract_list_value(dic):
    for key in dic.keys():
        dic[key] = dic[key][0]
    return dic

def write_nested_dicts_to_excel(WB, name, sheet, dic, range_list, header):
    if len(range_list) == 1:
        range1, range2 = calculate_write_ranges_dic(dic, range_list[0])
    else:
        range1 = range_list[0]
        range2 = range_list[1]
    multiple_cells = sheet[range1:range2]

    primary_keys = list(dic.keys())
    secondary_keys = list(dic[list(dic.keys())[0]].keys())

    primary_keys.sort(key=natural_keys)

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
    start_col = extract_letters(r1)
    start_row = extract_numbers(r1)
    r2 = chr(ord(start_col)+l2)+str(int(start_row)+l1)
    return r1, r2

def calculate_write_ranges_lst(lst, start):
    l1 = len(list(lst))
    r1 = start
    start_col = extract_letters(r1)
    start_row = extract_numbers(r1)
    r2 = start_col+str(int(start_row)+l1)
    return r1, r2

def calculate_write_ranges_from_dic_array(lst, start = 'D1'):
    # Se van a calcular los rangos de escritura de forma que queden de forma cuadrangular
    max_tables_in_row = int(sqrt(len(lst)))

    ranges = []

    new_start = start
    max_dic_size_in_row = 0
    permanent_increment = 0
    for i in range(len(lst)):
        dic = lst[i]
        new_range = calculate_write_ranges_dic(dic, new_start)
        ranges.append(new_range)
        if len(dic)>max_dic_size_in_row:
            max_dic_size_in_row = len(dic)
        if (i+1)%max_tables_in_row==0:
            permanent_increment += max_dic_size_in_row+2
            max_dic_size_in_row = 0
            new_start = extract_letters(start)+str(permanent_increment+int(extract_numbers(start)))
        else:
            new_start = calculate_new_range_start(new_range)
    return ranges
def calculate_new_range_start(ranges):
    r1 = ranges[0]
    r2 = ranges[1]
    start_row = extract_numbers(r1)
    end_col = extract_letters(r2)
    new_cell = chr(ord(end_col) + 2) + start_row
    return new_cell

def extract_letters(word):
    return re.findall(r'\D+', word)[0]

def extract_numbers(word):
    return re.findall(r'\d+', word)[0]

def extract_all_numbers(word):
    return re.findall(r'\d+', word)

def extract_complex_numbers(word):
    lst = re.findall(r'\d+', word)
    return int(''.join(lst))

def parse_complex_string_numbers(lst):
    return [extract_complex_numbers(word) for word in lst]


def extract_list_numbers(lst):
    lst_aux = []
    for item in lst:
        lst_aux.append(int(extract_numbers(item)))
    return lst_aux

def create_nested_dic(primary_key, secondary_key, value):
    aux = create_empty_nested_dics(primary_key)
    for i in primary_key:
        aux[i] = create_dic(secondary_key, value)
    return aux

def calculate_minimum(lst):
    m = lst[0]
    for i in lst:
        if i<m:
            m=i
    return m

def calculate_interval(string):
    a = string.split(' ')
    PM  = 'p.m.'
    AM  = 'a.m.'
    first_hour = int(a[0])
    second_hour = int(a[3])
    if a[1] == PM:
        first_hour += 12
    if a[4] == PM:
        second_hour += 12
    delay = second_hour - first_hour
    if delay < 0:
        delay += 24
    return delay

def calculate_interval_list(lst):
    lst_aux = []
    for item in lst:
        lst_aux.append(calculate_interval(item))
    return lst_aux

def atoi(text):
    return int(text) if text.isdigit() else text

def natural_keys(text):
    return [atoi(c) for c in re.split(r'(\d+)', text)]
