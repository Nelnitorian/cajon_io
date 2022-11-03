# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

def parse_list(lst):
    aux = []
    for item in lst:
        aux.append(int(item))
    return aux

def main():
    """
    x:
        x1 = xa
        x2 = xb
        ...

    v:
        v1 =  infraestructuras
        v2 = administración
        v3 = sector primario
        v4 = i+d
        v5 = desempleo

    :return:
    """
    name='NOMBRE_EXCEL.xlsx'
    excel_doc=openpyxl.load_workbook(name,data_only=True)
    sheet=excel_doc['Hoja 1']

    s = parse_list(Read_Excel_to_List(sheet, 'b2', 'b6'))
    p = parse_list(Read_Excel_to_List(sheet, 'c2', 'c7'))
    ni = parse_list(Read_Excel_to_List(sheet, 'd2', 'd6'))

    solver=pywraplp.Solver.CreateSolver('GLOP')

    x = []
    for i in range(len(s)):
        x.append(solver.NumVar(0, solver.infinity(), f"x{i}"))

    v = []
    for i in range(5):
        v.append(solver.NumVar(0, solver.infinity(), f"v{i}"))

    # R1
    solver.Add(sum(v)+7500+12500 >= 94500, f"R1")
    # R2
    solver.Add(sum(x)+7500+12500 >= 94500, f"R2")

    for i in fabricas:
        solver.Add(sum(x[i][j] for j in almacenes)==a[i-1], f"RF{i}")
    for j in almacenes:
        solver.Add(sum(x[i][j] for i in fabricas)==b[j-1], f"RA{j}")

    solver.Minimize(solver.Sum(s[i]*p[i]*(ni[i]+x[i]) for i in range(len(s))))

    status=solver.Solve()
    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion.')

        sol = {}
        crel = {}
        for i in fabricas:
            sol[i] = {j: 0.0 for j in almacenes}
            crel[i] = {j: 0.0 for j in almacenes}
        for i in fabricas:
            for j in almacenes:
                sol[i][j]=x[i][j].solution_value()
                crel[i][j]=x[i][j].reduced_cost()

        Write_NesteDic_to_Excel(excel_doc, name, sheet, sol,'f10', 'k16')
        Write_NesteDic_to_Excel(excel_doc, name, sheet, crel,'f18', 'k24')

    else:
        print('No hay solución óptima. Error.')


if __name__=='__main__':
    main()
