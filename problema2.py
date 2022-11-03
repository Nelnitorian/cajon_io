# -*- coding: utf-8 -*-

import openpyxl as openpyxl
from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

def parse_list(lst):
    aux = []
    for item in lst:
        if item:
            aux.append(int(item))
        else:
            aux.append(0)
    return aux

def main():
    name='trabajo.xlsx'
    excel_doc=openpyxl.load_workbook(name,data_only=True)
    sheet=excel_doc['Hoja 1']

    arcos = Read_Excel_to_List(sheet, 'a2', 'a11')
    nodos = parse_list(Read_Excel_to_List(sheet, 'f2', 'f7'))
    prod_dem = parse_list(Read_Excel_to_List(sheet, 'g2', 'g7'))
    info_arcos = Read_Excel_to_NesteDic(sheet, 'a1', 'd11')

    A = {}
    for n in nodos:
        A[n] = {}
        for a in arcos:
            A[n][a] = 0.0
    for n in nodos:
        for a in arcos:
            if info_arcos[a]['i'] == n:
                A[n][a] = 1
            elif info_arcos[a]['j'] == n:
                A[n][a] = -1
            else:
                A[n][a] = 0

    solver=pywraplp.Solver.CreateSolver('GLOP')

    x = {}

    for a in arcos:
        x[a] = solver.NumVar(0, solver.infinity(), f"X{a}")

    for n in nodos:
        solver.Add(sum(A[n][a]*x[a] for a in arcos) == prod_dem[n-1], f"RN{n}")

    solver.Minimize(sum(info_arcos[a]['coste']*x[a] for a in arcos))

    status=solver.Solve()
    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion')

        sol = {}
        crel = {}
        for a in arcos:
            sol[a] = x[a].solution_value()

        Write_DicTable_to_Excel(excel_doc, name, sheet, sol, 'j2', 'j11')


if __name__=='__main__':
    main()
