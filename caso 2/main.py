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

    # p = {}
    # for j in fabricantes:
    #     # Dado que python hace evaluación perezosa de las lists comprehensions es necesaria la siguiente sentencia de cara a la primera iteración del bucle:
    #     m = -1
    #     p[j] = solver.Sum(solver.Sum(x1[i][j] for i in proveedores if m in list(table_contents['T1'][table_contents['T1']["Proveedor"] == i]["Materia prima"]))*e[j][m] for m in materias)
    cantidad_materias = {}
    for j in fabricantes:
        aux = []
        for m in materias:
            cantidad = 0
            for i in proveedores:
                if m in list(table_contents['T1'][table_contents['T1']["Proveedor"] == i]["Materia prima"]):
                    cantidad += x1[i][j]
            aux.append(cantidad)
        cantidad_materias[j] = create_dic(materias, aux)
    p = create_empty_nested_dics(fabricantes)
    for j in fabricantes:
        for m in materias:
            p[j][m] = cantidad_materias[j][m]*e[j][m]
    p_total = {}
    for j in fabricantes:
        total = 0
        for m in materias:
            total += p[j][m]
        p_total[j] = total
    # print("a")
    # p = {}
    # for j in fabricantes:
    #     lst = []
    #     for m in materias:
    #         lst.append(solver.Sum(x1[i][j] for i in proveedores if m in list(table_contents['T1'][table_contents['T1']["Proveedor"] == i]["Materia prima"])) * e[j][m])
    #     p[j] = solver.Sum(lst)
    # materia1 = [x1[i][j] for j in fabricantes for i in [1,2,3]]
    # materia2 = [x1[i][j] for j in fabricantes for i in [4,5,6]]
    # materia3 = [x1[i][j] for j in fabricantes for i in [7,8]]
    # indexes = [[1,2,3],[4,5,6],[7,8]]
    # mats = []
    # for j in fabricantes:
    #     for m in materias:
    #         var_lst = []
    #         for j in fabricantes:
    #             tmp = []
    #             for lst in indexes:
    #                 for i in lst:
    #                     tmp.append(x1[i][j])
    #             var_lst.append(tmp)
    #         solver.Add(solver.Sum(var_lst)/e[j][m])
    #     solver.Add(solver.Sum([x1[i][j] for lst in indexes for j in fabricantes])/e[j][m])
    # solver.Add(solver.Sum(x1[i][j] for j in fabricantes for i in [4,5,6]))
    # solver.Add(solver.Sum(x1[i][j] for j in fabricantes for i in [7,8]))

        # R1
    for i in proveedores:
        solver.Add(solver.Sum(x1[i][j] for j in fabricantes) <= a1[i], f"R1")

    # R2
    for j in fabricantes:
        solver.Add(solver.Sum(x1[i][j] for i in proveedores) == b1[j], f"R2")

    # R3
    for i in fabricantes:
        solver.Add(solver.Sum(x2[i][j] for j in clientes) <= p_total[i], f"R3")

    # R4
    for j in clientes:
        solver.Add(solver.Sum(x2[i][j] for i in fabricantes) == b2[j], f"R4")

    # R5
    for i in fabricantes:
        solver.Add(p_total[i] <= a2[i], f"R5")

    FO = solver.Sum(x1[i][j]*(kf1[i]+kd1[i]*d1[i][j]) for j in fabricantes for i in proveedores) + solver.Sum(x2[i][j]*(kf2[i]+kd2[i][j]*d2[i][j]) for j in clientes for i in fabricantes)
    solver.Minimize(FO)

    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')


        # Cantidad de producto enviada entre proveedores y fabricantes
        sol_x1 = create_empty_nested_dics(proveedores)
        for i in proveedores:
            for j in fabricantes:
                sol_x1[i][j] = x1[i][j].solution_value()

        sol_x2 = create_empty_nested_dics(fabricantes)
        for i in fabricantes:
            for j in clientes:
                sol_x2[i][j] = x2[i][j].solution_value()

        var_costs_x1 = create_empty_nested_dics(proveedores)
        for i in proveedores:
            for j in fabricantes:
                var_costs_x1[i][j] = x1[i][j].reduced_cost()

        var_costs_x2 = create_empty_nested_dics(fabricantes)
        for i in fabricantes:
            for j in clientes:
                var_costs_x2[i][j] = x2[i][j].reduced_cost()

        FO_value = sum(sol_x1[i][j]*(kf1[i]+kd1[i]*d1[i][j]) for j in fabricantes for i in proveedores) + sum(sol_x2[i][j]*(kf2[i]+kd2[i][j]*d2[i][j]) for j in clientes for i in fabricantes)

        transport_money_costs_x1 = create_empty_nested_dics(proveedores)
        for i in proveedores:
            for j in fabricantes:
                transport_money_costs_x1[i][j] = sol_x1[i][j]*kd1[i]*d1[i][j]

        transport_money_costs_x2 = create_empty_nested_dics(fabricantes)
        for i in fabricantes:
            for j in clientes:
                transport_money_costs_x2[i][j] = sol_x2[i][j]*kd2[i][j]*d2[i][j]

        fabrication_money_costs_x1 = create_empty_nested_dics(proveedores)
        for i in proveedores:
            for j in fabricantes:
                fabrication_money_costs_x1[i][j] = sol_x1[i][j]*kf1[i]

        fabrication_money_costs_x2 = create_empty_nested_dics(fabricantes)
        for i in fabricantes:
            for j in clientes:
                fabrication_money_costs_x2[i][j] = sol_x2[i][j]*kf2[i]

        total_money_cost_x1 = create_empty_nested_dics(proveedores)
        for i in proveedores:
            for j in fabricantes:
                total_money_cost_x1[i][j] = fabrication_money_costs_x1[i][j]+transport_money_costs_x1[i][j]

        total_money_cost_x2 = create_empty_nested_dics(fabricantes)
        for i in fabricantes:
            for j in clientes:
                total_money_cost_x2[i][j] = fabrication_money_costs_x2[i][j]+transport_money_costs_x2[i][j]

        product_by_mat = create_empty_nested_dics(fabricantes)
        for j in fabricantes:
            for m in materias:
                product_by_mat[j][m] = p[j][m].solution_value()

        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, sol_x1, ['E28'], 'S1')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, sol_x2, ['I28'], 'S2')
        # write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, var_costs_x1, ['N28'],'S3')
        # write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, var_costs_x2, ['E38'], 'S4')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, transport_money_costs_x1, ['I38'], 'S5')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, transport_money_costs_x2, ['N38'], 'S6')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, fabrication_money_costs_x1,['E48'] , 'S7')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, fabrication_money_costs_x2, ['I48'], 'S8')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, total_money_cost_x1, ['N48'], 'S9')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, total_money_cost_x2, ['E58'], 'S10')
        write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, product_by_mat, ['I58'], 'S11')
        # funcion objetivo
        write_list_to_excel(excel_doc, EXCEL_FILE_NAME, sheet, [FO_value,], ['A28'], 'Valor de la Función Objetivo')

    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
