# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from utils import *
import pandas as pd

# Constants
EXCEL_FILE_NAME = 'caso4_excel.xlsx'
DATA_SHEET_NAME = 'Datos'
ANSWER_SHEET_NAME = 'Resultados'


def main():
    """


    :return:
    """


    solver = pywraplp.Solver.CreateSolver('SCIP')

    # Read from excel
    excel_doc = openpyxl.load_workbook(EXCEL_FILE_NAME, data_only=True)
    data_sheet = excel_doc[DATA_SHEET_NAME]

    # Se lee el diccionario donde está la información de las tablas
    table_information = parse_dic_values(read_dic_single_value_from_excel(data_sheet, ('A2', 'B5')))
    table_contents = {}
    for key, value in table_information.items():
        table_contents[key] = read_from_excel_to_dataframe(EXCEL_FILE_NAME, DATA_SHEET_NAME, value)



    # Definición de indices

    tipos = list(table_contents['T2']["Datos producción"])
    intervalos = list(table_contents['T1']['ID'])

    # Definición de constantes
    CANTIDAD_POR_TIPO = create_dic(tipos, [list(range(12)), list(range(10)), list(range(5))])
    d = create_dic(tipos, calculate_interval_list(table_contents['T1']['Hora']))
    n = create_dic(tipos, table_contents['T1']['Demanda'])
    pmin = create_dic(tipos, extract_list_numbers(table_contents['T2']['Mínima producción']))
    pmax = create_dic(tipos, extract_list_numbers(table_contents['T2']['Mínima producción']))
    cmin = create_dic(tipos, table_contents['T2']['Coste por hora a mínimo nivel'])
    cext = create_dic(tipos, table_contents['T2']['Coste extra por MW sobre el mínimo'])
    capt = create_dic(tipos, table_contents['T2']['Coste de arranque'])


    p = create_empty_nested_dics(intervalos)
    for i in intervalos:
        p[i] = create_empty_nested_dics(tipos)
    for i in intervalos:
        for t in tipos:
            for n in CANTIDAD_POR_TIPO[t]:
                p[i][t][n] = solver.IntVar(pmin[t], pmax[t], f"p_{i}_{t}_{n}")

    o = create_empty_nested_dics(intervalos)
    for i in intervalos:
        o[i] = create_empty_nested_dics(tipos)
    for i in intervalos:
        for t in tipos:
            for n in CANTIDAD_POR_TIPO[t]:
                o[i][t][n] = solver.NumVar(0, 1, f"o_{i}_{t}_{n}")

    # R3
    for i in intervalos:
        solver.Add(solver.Sum(p[i][t][n]*o[i][t][n] for t in tipos for n in CANTIDAD_POR_TIPO[t]) >= n[i], f"R3_{i}")

    # R4
    for i in intervalos:
        solver.Add(solver.Sum(pmax[t]*o[i][t][n] for t in tipos for n in CANTIDAD_POR_TIPO[t]) >= 1.15*n[i], f"R4_{i}")

    FO = solver.Sum(d[i]*((p[i][t][n]-pmin[t])*cext[t]+cmin[t])*o[i][t][n] for i in intervalos for t in tipos for n in CANTIDAD_POR_TIPO[t]) +\
        solver.Sum(o[intervalos[0]][t][n]*capt[t] for t in tipos for n in CANTIDAD_POR_TIPO[t]) +\
        solver.Sum((1-o[intervalos[i]][t][n]*capt[t])*o[intervalos[i+1]][t][n]*capt[t] for i in range(len(intervalos)-1) for t in tipos for n in CANTIDAD_POR_TIPO[t])
    solver.Minimize(FO)
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        p_sol = create_empty_nested_dics(intervalos)
        for i in intervalos:
            p_sol[i] = create_empty_nested_dics(tipos)
        for i in intervalos:
            for t in tipos:
                for n in CANTIDAD_POR_TIPO[t]:
                    p_sol[i][t][n] = p[i][t][n].solution_value()

        o_sol = create_empty_nested_dics(intervalos)
        for i in intervalos:
            o_sol[i] = create_empty_nested_dics(tipos)
        for i in intervalos:
            for t in tipos:
                for n in CANTIDAD_POR_TIPO[t]:
                    o_sol[i][t][n] = o[i][t][n].solution_value()

        suma_modulos_abiertos = create_empty_nested_dics(intervalos)
        for i in intervalos:
            for t in tipos:
                suma_modulos_abiertos[i][t] = sum(o_sol[i][t])

        total_marginal_cost = create_empty_nested_dics(intervalos)
        for i in intervalos:
            i_FO_value = sum(
                d[i] * ((p_sol[i][t][n] - pmin[t]) * cext[t] + cmin[t]) * o_sol[i][t][n] for i in intervalos for t in tipos for
                n in CANTIDAD_POR_TIPO[t]) + \
                 sum(o_sol[intervalos[0]][t][n] * capt[t] for t in tipos for n in CANTIDAD_POR_TIPO[t]) + \
                 sum((1 - o_sol[intervalos[i]][t][n] * capt[t]) * o_sol[intervalos[i + 1]][t][n] * capt[t] for i in
                            range(len(intervalos) - 1) for t in tipos for n in CANTIDAD_POR_TIPO[t])
            total_marginal_cost[i]["Coste"] = i_FO_value

        total_marginal_production = create_empty_nested_dics(intervalos)
        for i in intervalos:
            total_marginal_production[i]["A"] = sum(p_sol[i][t][n] for t in tipos for n in CANTIDAD_POR_TIPO[t])

        marginal_cost = create_empty_nested_dics(intervalos)
        for i in intervalos:
            marginal_cost[i]["Precio mínimo del MW"] = total_marginal_cost[i]["Coste"]/total_marginal_production[i][""]

        answer_sheet = excel_doc[ANSWER_SHEET_NAME]
        dics_array = [suma_modulos_abiertos, total_marginal_cost, total_marginal_production,
                      marginal_cost]
        ranges = calculate_write_ranges_from_dic_array(dics_array, start='C3')
        for i in range(len(dics_array)):
            write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, answer_sheet, dics_array[i], ranges[i], f'S{i+1}')

        print(f"Valor de la función objetivo total: {FO.solution_value()}")
    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
