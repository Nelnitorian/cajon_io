# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *
import pandas as pd

# Constants
EXCEL_FILE_NAME = 'caso2_excel.xlsx'
SHEET_NAME ='Hoja 1'

def main():
    """


    :return:
    """


    solver = pywraplp.Solver.CreateSolver('GLOP')

    # Read from excel
    excel_doc=openpyxl.load_workbook(EXCEL_FILE_NAME, data_only=True)
    sheet=excel_doc[SHEET_NAME]

    # Se lee el diccionario donde está la información de las tablas
    table_information = parse_dic_values(Read_Dic_Single_Value_from_Excel(sheet, ('A2', 'B8')))
    table_contents = {}
    for key, value in table_information.items():
        table_contents[key] = read_from_excel_to_dataframe(EXCEL_FILE_NAME, SHEET_NAME, value)

    # Definición de variables

    x1 = create_dictionary_vars(solver, list(table_contents['T1']['Proveedor']), list(table_contents['T4']['Fábrica']))
    x2 = create_dictionary_vars(solver, list(table_contents['T4']['Fábrica']), list(table_contents['T5'].keys()))

    # Definición de constantes. Se pasan a diccionarios para mejorar la trazabilidad.

    a1 = create_dic(x1.keys(), table_contents['T1']["Límite de suministro (unidades/año)"])
    a2 = create_dic(x2.keys(), table_contents['T4']["Unidades de producto/año"])

    kf1 = create_dic(x1.keys(), table_contents['T1']["Coste unitario"])
    kf2 = create_dic(x2.keys(), table_contents['T4']["Coste de fabricación"])



    kd1 = create_dic(x1.keys(), table_contents['T4']["Coste de fabricación"])
    kd2 = create_dic(x2.keys(), table_contents['T4']["Coste de fabricación"])

    d1 = create_dic(x1.keys(), table_contents['T4']["Coste de fabricación"])
    d2 = create_dic(x2.keys(), table_contents['T4']["Coste de fabricación"])

    # Restricciones


    # R1
    solver.Add(solver.Sum(v[i] for i in range(len(v))) >= 0.75*g['TOTAL'], f"R1")
    # R2
    solver.Add((solver.Sum(s[i]*p[i]*(ni[i]+x[i]) for i in range(len(s))) -
               solver.Sum(v[i] for i in range(len(v)))) == 27000, f"R2")
    # R3
    solver.Add(v[4] >= 2500, f"R3")
    # R4
    solver.Add(v[5] >= 900, f"R4")
    # R5
    solver.Add(v[6] >= 6000, f"R5")
    # R6
    solver.Add(x[0]/s[0] == x[1]/s[1], f"R6")
    # R7
    solver.Add(x[1]/s[1] == x[2]/s[2], f"R7")
    # R8
    solver.Add(x[2]/s[2] == x[3]/s[3], f"R8")
    # R9
    solver.Add(x[3]/s[3] == x[4]/s[4], f"R9")
    # R10
    solver.Add(v[0]/g['INFRA'] == v[3]/g['ADMIN'], f"R10")
    # R11
    solver.Add(v[1] == g['EDUCACION'], f"R11")
    # R12
    solver.Add(v[2] == g['SANIDAD'], f"R12")

    solver.Minimize(solver.Sum(s[i]*p[i]*(ni[i]+x[i]) for i in range(len(s))))

    status=solver.Solve()

    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        t = []
        for sol in x:
            t.append(sol.solution_value())
        n_ni = [ni[i] + t[i] for i in range(len(ni))]
        ci = []
        for cii in v:
            ci.append(cii.solution_value())
        ing = sum([s[i]*p[i]*(ni[i]+t[i]) for i in range(len(s))])
        gas = sum(ci)
        print(f"El incremento de las tasas impositivas es: {t}")
        print(f"Dejando las tasas impositivas como: {n_ni}")
        print(f"Los nuevos costes de inversión son: {ci}")
        print(f"Los ingresos son: {round(ing, 2)}")
        print(f"Los gastos son: {gas}")
        print(f"Balance del año: {round(ing-gas, 2)}")


    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
