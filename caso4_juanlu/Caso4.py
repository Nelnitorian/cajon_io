# -*- coding: utf-8 -*-
"""

@author: Juan Luis Verdugo Blanco
         Aitor Laiseca Valencia
         Mario Sáenz de Ormijana Nieva
"""

from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

name='Datos.xlsx'
excel_doc=openpyxl.load_workbook(name,data_only=True)
sheet=excel_doc['Hoja1']


h_tramos=Read_Excel_to_List(sheet, 'e4', 'e8')          #Horas por tramo horario
demanda=Read_Excel_to_List(sheet, 'c4', 'c8')         #MW por hora en cada tramo
min_prod=Read_Excel_to_List(sheet, 'c13', 'c15')      #Mínima producción de cada tipo
max_prod=Read_Excel_to_List(sheet, 'd13', 'd15')      #Máxima producción de cada tipo
c_hora_min=Read_Excel_to_List(sheet, 'e13', 'e15')    #Coste por hora a nivel mínimo
c_extra=Read_Excel_to_List(sheet, 'f13', 'f15')       #Coste extra por MW sobre mínimo
c_arranque=Read_Excel_to_List(sheet, 'g13', 'g15')    #Coste arranque
unidades_gen=Read_Excel_to_List(sheet, 'c19', 'c21')  #Número unidades generación por tipo
prod_reserva=Read_Excel_to_List(sheet, 'b24','b24')   #Porcentaje prod reserva 

Tramos=Read_Excel_to_List(sheet, 'a4', 'a8')
Tipo1=[n+1 for n in range(unidades_gen[0])]
Tipo2=[n+1 for n in range(unidades_gen[1])]
Tipo3=[n+1 for n in range(unidades_gen[2])]


demanda_h=[];
for t in Tramos:
    demanda_h.insert(t-1, h_tramos[t-1]*demanda[t-1]);

def Problema():
    
    solver=pywraplp.Solver.CreateSolver('CBC')
    
    
    x={}
    for i in Tipo1:
        x[i]={}
        for t in Tramos:
            x[i][t]=solver.IntVar(0, solver.infinity(), 'X%d,%d'%(i,t))

    y={}
    for j in Tipo2:
        y[j]={}
        for t in Tramos:
            y[j][t]=solver.IntVar(0, solver.infinity(), 'Y%d,%d'%(j,t))
            
    z={}
    for k in Tipo3:
        z[k]={}
        for t in Tramos:
            z[k][t]=solver.IntVar(0, solver.infinity(), 'Z%d,%d'%(k,t))
            
    deltax={}
    for i in Tipo1:
        deltax[i]={}
        for t in Tramos:
            deltax[i][t]=solver.IntVar(0, 1, 'deltaX%d,%d'%(i,t))
            
    deltay={}
    for j in Tipo2:
        deltay[j]={}
        for t in Tramos:
            deltay[j][t]=solver.IntVar(0, 1, 'deltaY%d,%d'%(j,t))
            
            
    deltaz={}
    for k in Tipo3:
        deltaz[k]={}
        for t in Tramos:
            deltaz[k][t]=solver.IntVar(0, 1, 'deltaZ%d,%d'%(k,t))
    
              
  
    for t in Tramos:
        for i in Tipo1:
            solver.Add(x[i][t]>= min_prod[0]*deltax[i][t])
            
    for t in Tramos:
        for j in Tipo2:
            solver.Add(y[j][t]>= min_prod[1]*deltay[j][t])
            
    for t in Tramos:
        for k in Tipo3:
            solver.Add(z[k][t]>= min_prod[2]*deltaz[k][t])
                        
    for t in Tramos:
        for i in Tipo1:
            solver.Add(x[i][t]<= max_prod[0]*deltax[i][t])
            
    for t in Tramos:
        for j in Tipo2:
            solver.Add(y[j][t]<= max_prod[1]*deltay[j][t])
            
    for t in Tramos:
        for k in Tipo3:
            solver.Add(z[k][t]<= max_prod[2]*deltaz[k][t])
                        
    for t in Tramos:
        solver.Add(solver.Sum(x[i][t] for i in Tipo1)+solver.Sum(y[j][t] for j in Tipo2)+solver.Sum(z[k][t] for k in Tipo3) >= demanda_h[t-1], 'RF%d'%(t))
                   
        
    for t in Tramos:
        solver.Add(sum(max_prod[0]-x[i][t] for i in Tipo1)+sum(max_prod[1]-y[j][t] for j in Tipo2)+sum(max_prod[2]-z[k][t] for k in Tipo3) >= demanda_h[t-1]*prod_reserva[0], 'RT%d'%(t))

      
    solver.Minimize(solver.Sum(((x[i][t]-min_prod[0]*deltax[i][t])*c_extra[0]) + c_hora_min[0]*h_tramos[0]*deltax[i][t] for i in Tipo1 for t in Tramos)
                    +solver.Sum(((y[j][t]-min_prod[1]*deltay[j][t])*c_extra[1]) + c_hora_min[1]*h_tramos[1]*deltay[j][t]  for j in Tipo2 for t in Tramos)
                    +solver.Sum(((z[k][t]-min_prod[2]*deltaz[k][t])*c_extra[2]) + c_hora_min[2]*h_tramos[2]*deltaz[k][t]  for k in Tipo3 for t in Tramos))

    status=solver.Solve()
    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion')
        objective = solver.Objective()
        print('El valor total de los costes según la minimización de la función objetivo es: {:.2f}' . format (solver.Objective().Value()))
        for i in Tipo1:
            for t in Tramos:
                print ('Valor x es: {:.2f}' . format (x[i][t].solution_value()))                
        for j in Tipo2:
            for t in Tramos:
                print ('Valor y es: {:.2f}' . format (y[j][t].solution_value()))
        for k in Tipo3:
            for t in Tramos:
                print ('Valor z es: {:.2f}' . format (z[k][t].solution_value()))
                
        for i in Tipo1:
            for t in Tramos:
                print ('Valor deltax es: {:.2f}' . format (deltax[i][t].solution_value()))
        for j in Tipo2:
            for t in Tramos:
                print ('Valor deltay es: {:.2f}' . format (deltay[j][t].solution_value()))
        for k in Tipo3:
            for t in Tramos:
                print ('Valor deltaz es: {:.2f}' . format (deltaz[k][t].solution_value()))
    else:
        print("Chiquitriquis")


Problema()