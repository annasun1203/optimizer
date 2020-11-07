#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from gurobipy import Model, GRB
def transportationOptimizer(inputFile,outputFile):
    ''' the function takes in two input parameters: 
        inputFile: file name of the input file
        outputFile: file name of the output file, which is created by the function.
    The code uses Gurobi to assign salespeople to specific conventions while minimizing the total 
    transportation costs.
    ''' 
    c=pd.read_excel(inputFile,sheet_name='costs',index_col=0)
    s=pd.read_excel(inputFile,sheet_name='supply',index_col=0)['available']
    d=pd.read_excel(inputFile,sheet_name='demand',index_col=0)['needed']
    mod=Model()
    I=c.index
    J=c.columns
    x=mod.addVars(I,J,vtype=GRB.INTEGER)
    mod.setObjective(sum(x[i,j]*c.loc[i,j] for i in I for j in J))
    for i in I:
        mod.addConstr(sum(x[i,j] for j in J)<=s.loc[i])
    for j in J:
        mod.addConstr(sum(x[i,j] for i in I)>=d.loc[j])
    mod.setParam('outputflag',False)
    mod.optimize()
    
    # Write your code for creating the output file below
    writer = pd.ExcelWriter('12-transportation-sampleOutput-1.xlsx')
    pd.DataFrame([mod.objval], columns = ['Minimumal cost']).to_excel(writer,sheet_name = 'Objective', index = False)
    df = pd.DataFrame(index = I, columns = J)
    for i in I:
        for j in J:
            df.loc[i,j]=x[i,j].x
    df.to_excel(writer, sheet_name = 'Plan') # second sheet 
    writer.save()
    
    





