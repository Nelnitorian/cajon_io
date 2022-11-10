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
    table_information = parse_dic_values(Read_Dic_Single_Value_from_Excel(sheet, ('A2', 'B8')), ',')
    table_contents = {}
    for key, value in table_information.items():
        table_contents[key] = read_from_excel_to_dataframe(EXCEL_FILE_NAME, SHEET_NAME, value)

    # Definición de variables

    x1 = create_dictionary_vars(solver, list(table_contents['T1']['Proveedor']), list(table_contents['T4']['Fábrica']))
    x2 = create_dictionary_vars(solver, list(table_contents['T4']['Fábrica']), list(table_contents['T5'].keys()))

    proveedores = list(x1.keys())
    fabricantes = list(x2.keys())
    clientes = list(table_contents['T5'].keys())
    materias = list(table_contents['T1']["Materia prima"].unique())

    # Definición de constantes. Se pasan a diccionarios para mejorar la trazabilidad.

    a1 = create_dic(proveedores, table_contents['T1']["Límite de suministro (unidades/año)"])
    a2 = create_dic(fabricantes, table_contents['T4']["Unidades de producto/año"])

    kf1 = create_dic(proveedores, table_contents['T1']["Coste unitario"])
    kf2 = create_dic(fabricantes, table_contents['T4']["Coste de fabricación"])

    kd1 = create_dic(proveedores, table_contents['T3']['Coste unitario de transporte [€/(unidad*km)]'])
    kd2 = {}
    for i in range(len(fabricantes)):
        kd2[fabricantes[i]] = create_dic(clientes, list(table_contents['T7'].iloc[0, 1:]))

    d1 = {}
    for i in range(len(proveedores)):
        d1[proveedores[i]] = create_dic(fabricantes, list(table_contents['T3'].iloc[i, 2:4]))
    d2 = {}
    for i in range(len(fabricantes)):
        d2[fabricantes[i]] = create_dic(clientes, list(table_contents['T6'].iloc[i, 1:]))


    e = {}
    for i in range(len(fabricantes)):
        e[fabricantes[i]] = create_dic(materias, list(table_contents['T2'].iloc[:, i+1]))

    b1 = {}
    for j in fabricantes:
        b1[j] = solver.Sum(x1[i][j] for i in proveedores)
    b2 = create_dic(clientes, list(table_contents['T5'].iloc[0, :]))

    p = {}
    for j in fabricantes:
        # Dado que python hace evaluación perezosa de las lists comprehensions es necesaria la siguiente sentencia de cara a la primera iteración del bucle:
        m = -1
        p[j] = solver.Sum(solver.Sum(x1[i][j] for i in proveedores if m in list(table_contents['T1'][table_contents['T1']["Proveedor"] == i]["Materia prima"]))*e[j][m] for m in materias)

    # Restricciones

    # R1
    for i in proveedores:
        solver.Add(solver.Sum(x1[i][j] for j in fabricantes) <= a1[i], f"R1")

    # R2
    for j in fabricantes:
        solver.Add(solver.Sum(x1[i][j] for i in proveedores) == b1[j], f"R2")

    # R3
    for i in fabricantes:
        solver.Add(solver.Sum(x2[i][j] for j in clientes) <= p[i], f"R3")

    # R4
    for j in clientes:
        solver.Add(solver.Sum(x2[i][j] for i in fabricantes) == b2[j], f"R4")

    # R5
    for i in fabricantes:
        solver.Add(p[i] <= a2[i], f"R5")

    FO = solver.Sum(x1[i][j]*(kf1[i]+kd1[i]) for j in fabricantes for i in proveedores) + solver.Sum(x2[i][j]*(kf2[i]+kd2[i][j]) for j in clientes for i in fabricantes)
    solver.Minimize(FO)

    status=solver.Solve()

    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        t = []
        # for sol in x:
        #     t.append(sol.solution_value())
        # n_ni = [ni[i] + t[i] for i in range(len(ni))]
        # ci = []
        # for cii in v:
        #     ci.append(cii.solution_value())
        # ing = sum([s[i]*p[i]*(ni[i]+t[i]) for i in range(len(s))])
        # gas = sum(ci)
        # print(f"El incremento de las tasas impositivas es: {t}")
        # print(f"Dejando las tasas impositivas como: {n_ni}")
        # print(f"Los nuevos costes de inversión son: {ci}")
        # print(f"Los ingresos son: {round(ing, 2)}")
        # print(f"Los gastos son: {gas}")
        # print(f"Balance del año: {round(ing-gas, 2)}")


    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
