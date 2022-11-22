# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from utils import *
import pandas as pd

# Constants
EXCEL_FILE_NAME = 'caso3_excel.xlsx'
DATA_SHEET_NAME = 'Datos'
ANSWER_SHEET_BASE_NAME = 'Anio'
N_INITIAL_FACTORIES = 2

def main():
    """


    :return:
    """


    solver = pywraplp.Solver.CreateSolver('GLOP')

    # Read from excel
    excel_doc = openpyxl.load_workbook(EXCEL_FILE_NAME, data_only=True)
    data_sheet = excel_doc[DATA_SHEET_NAME]

    # Se lee el diccionario donde está la información de las tablas
    table_information = parse_dic_values(read_dic_single_value_from_excel(data_sheet, ('A2', 'B9')))
    table_contents = {}
    for key, value in table_information.items():
        table_contents[key] = read_from_excel_to_dataframe(EXCEL_FILE_NAME, DATA_SHEET_NAME, value)

    # Definición de variables

    proveedores = list(table_contents['T1']['Proveedor'])
    fabricantes = list(table_contents['T4']['Fábrica'])
    clientes = list(table_contents['T5'].keys()[1:])
    materias = list(table_contents['T1']["Materia prima"].unique())
    anios = list(table_contents['T5']["Demanda/Anios"])

    x1 = create_empty_nested_dics(anios)
    for t in anios:
        x1[t] = create_dictionary_vars(solver, proveedores, fabricantes)
    # x1 = create_dictionary_vars(solver, proveedores, fabricantes)
    x2 = create_empty_nested_dics(anios)
    for t in anios:
        x2[t] = create_dictionary_vars(solver, fabricantes, clientes)

    # Inicialización de la variable binaria de operación
    o = create_empty_nested_dics(anios)
    for t in range(len(anios)-1):
        for i in fabricantes:
            o[anios[t+1]][i] = solver.NumVar(0, 1, f"o_{anios[t+1]}_{i}")
    # Inicialización del primer año de operación
    for i in range(len(fabricantes)):
        if i<N_INITIAL_FACTORIES:
            o[anios[0]][fabricantes[i]] = 1
        else:
            o[anios[0]][fabricantes[i]] = 0

    # Definición de constantes. Se pasan a diccionarios para mejorar la trazabilidad.

    a1 = create_dic(proveedores, table_contents['T1']["Límite de suministro (unidades/año)"])
    a2 = create_dic(fabricantes, table_contents['T4']["Unidades de producto/año"])

    ka = create_dic(fabricantes, table_contents['T4']["Coste fijo anual"]) #coste de operacion
    G = create_dic(fabricantes, table_contents['T8']["Beneficio por cierre"]) #coste por cierre (beneficios)
    I = create_dic(fabricantes, table_contents['T8']["Inversión de apertura"]) #coste por apertura (inversion)

    kf1 = create_dic(proveedores, table_contents['T1']["Coste unitario"])
    kf2 = create_dic(fabricantes, table_contents['T4']["Coste de fabricación"])

    kd1 = create_dic(proveedores, table_contents['T3']['Coste unitario de transporte [€/(unidad*km)]'])
    kd2 = {}
    for i in range(len(fabricantes)):
        kd2[fabricantes[i]] = create_dic(clientes, list(table_contents['T7'].iloc[0, 1:]))

    d1 = {}
    for i in range(len(proveedores)):
        d1[proveedores[i]] = create_dic(fabricantes, list(table_contents['T3'].iloc[i, 2:-1]))
    d2 = {}
    for i in range(len(fabricantes)):
        d2[fabricantes[i]] = create_dic(clientes, list(table_contents['T6'].iloc[i, 1:]))


    e = {}
    for i in range(len(fabricantes)):
        e[fabricantes[i]] = create_dic(materias, list(table_contents['T2'].iloc[:, i+1]))

    b1 = create_empty_nested_dics(anios)
    for t in anios:
        for j in fabricantes:
            b1[t][j] = solver.Sum(x1[t][i][j] for i in proveedores)
    b2 = create_empty_nested_dics(anios)
    for t in range(len(anios)):
        b2[anios[t]] = create_dic(clientes, list(table_contents['T5'].iloc[t, :])[1:])

    cantidad_materias = create_empty_nested_dics(anios)
    for t in anios:
        for j in fabricantes:
            aux = []
            for m in materias:
                cantidad = 0
                for i in proveedores:
                    if m in list(table_contents['T1'][table_contents['T1']["Proveedor"] == i]["Materia prima"]):
                        cantidad += x1[t][i][j]
                aux.append(cantidad)
            cantidad_materias[t][j] = create_dic(materias, aux)
    p = create_empty_nested_dics(anios)
    for t in anios:
        p[t] = create_empty_nested_dics(fabricantes)
    for t in anios:
        for j in fabricantes:
            for m in materias:
                p[t][j][m] = cantidad_materias[t][j][m]/e[j][m]
    p_total = create_empty_nested_dics(anios)
    for t in anios:
        for j in fabricantes:
            p_total[t][j] = p[t][j][materias[0]]

    # R1
    for t in anios:
        for i in proveedores:
            solver.Add(solver.Sum(x1[t][i][j] for j in fabricantes) <= a1[i], f"R1")

    # R2
    for t in anios:
        for j in fabricantes:
            solver.Add(solver.Sum(x1[t][i][j] if o[t][j] else 0 for i in proveedores) == b1[t][j], f"R2")

    # R3
    for t in anios:
        for i in fabricantes:
            solver.Add(solver.Sum(x2[t][i][j] if o[t][i] else 0 for j in clientes) <= p_total[t][i], f"R3")

    # R4
    for t in anios:
        for j in clientes:
            solver.Add(solver.Sum(x2[t][i][j] if o[t][i] else 0 for i in fabricantes) == b2[t][j], f"R4")

    # R5
    for t in anios:
        for i in fabricantes:
            solver.Add(p_total[t][i] <= a2[i], f"R5")

    # R6
    for t in range(len(anios)-1):
        for j in fabricantes:
            if j in fabricantes[:N_INITIAL_FACTORIES]:
                solver.Add(o[anios[t]][j] >= o[anios[t+1]][j], f"R6")
            else:
                solver.Add(o[anios[t]][j] <= o[anios[t+1]][j], f"R6")

    #R7
    for t in anios:
        for j in fabricantes:
            for m in materias:
                for m2 in materias:
                    solver.Add(p[t][j][m] == p[t][j][m2], f"R7")

    FO = solver.Sum(x1[t][i][j]*(kf1[i]+kd1[i]*d1[i][j]) if o[t][j] == 1 else 0 for j in fabricantes for i in proveedores for t in anios) + \
         solver.Sum(x2[t][i][j]*(kf2[i]+kd2[i][j]*d2[i][j]) if o[t][i] == 1 else 0 for j in clientes for i in fabricantes for t in anios) + \
         solver.Sum(ka[i]*o[t][i] for i in fabricantes for t in anios) + \
         -1*solver.Sum(G[fabricantes[i]]*(1-o[anios[-1]][fabricantes[i]]) for i in range(N_INITIAL_FACTORIES)) + \
         solver.Sum(I[fabricantes[i+N_INITIAL_FACTORIES]]*o[anios[-1]][fabricantes[i]] for i in range(len(fabricantes)-N_INITIAL_FACTORIES))
    solver.Minimize(FO)
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        operacion = create_empty_nested_dics(anios)
        for t in range(len(anios)-1):
            for j in fabricantes:
                operacion[anios[t+1]][j] = o[anios[t+1]][j].solution_value()
        for j in fabricantes:
            operacion[anios[0]][j] = o[anios[0]][j]

        for t in anios:
            # Cantidad de producto enviada entre proveedores y fabricantes
            sol_x1 = create_empty_nested_dics(proveedores)
            for i in proveedores:
                for j in fabricantes:
                    sol_x1[i][j] = x1[t][i][j].solution_value()

            sol_x2 = create_empty_nested_dics(fabricantes)
            for i in fabricantes:
                for j in clientes:
                    sol_x2[i][j] = x2[t][i][j].solution_value()

            var_costs_x1 = create_empty_nested_dics(proveedores)
            for i in proveedores:
                for j in fabricantes:
                    var_costs_x1[i][j] = x1[t][i][j].reduced_cost()

            var_costs_x2 = create_empty_nested_dics(fabricantes)
            for i in fabricantes:
                for j in clientes:
                    var_costs_x2[i][j] = x2[t][i][j].reduced_cost()

            t_FO_value = sum(sol_x1[i][j]*(kf1[i]+kd1[i]*d1[i][j]) for j in fabricantes for i in proveedores) + \
                         sum(sol_x2[i][j]*(kf2[i]+kd2[i][j]*d2[i][j]) for j in clientes for i in fabricantes) + \
                         sum(I[i] if not operacion[anios[t2] + 1][i] == operacion[anios[t2]][i] and operacion[anios[t2]][i] == 1 else 0 for i in fabricantes for t2 in range(len(anios) - 1)) + \
                         -1*sum(G[i] if not operacion[anios[t2] + 1][i] == operacion[anios[t2]][i] and operacion[anios[t2]][i] == 0 else 0 for i in fabricantes for t2 in range(len(anios) - 1)) + \
                         sum(ka[i]*o[t2][i] if operacion[t2][i] == 1 else 0 for i in fabricantes for t2 in anios)

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
                    product_by_mat[j][m] = p[t][j][m].solution_value()

            answer_sheet = excel_doc[ANSWER_SHEET_BASE_NAME+str(int(t))]
            dics_array = [sol_x1, sol_x2, transport_money_costs_x1,
                          transport_money_costs_x2, fabrication_money_costs_x1, fabrication_money_costs_x2,
                          total_money_cost_x1, total_money_cost_x2, product_by_mat, operacion]
            ranges = calculate_write_ranges_from_dic_array(dics_array, start='D1')
            for i in range(len(dics_array)):
                write_nested_dicts_to_excel(excel_doc, EXCEL_FILE_NAME, answer_sheet, dics_array[i], ranges[i], f'S{i+1}')

            # funcion objetivo
            write_list_to_excel(excel_doc, EXCEL_FILE_NAME, answer_sheet, [t_FO_value, ], ['A4'], 'Valor de la Función Objetivo')

    else:
        print('No hay solución óptima. Error.')

    return

if __name__=='__main__':
    main()
