# -*- coding: utf-8 -*-
"""
By: Peng Shi
Illustration code for Week 12.
"""

import pandas as pd
from gurobipy import Model, GRB
def transportationOptimizer(inputFile,outputFile):
    ''' The function takes in two arguments:
        inputFile: filename of the input file.
        outputFile: filename of the output file, which is created by this function.
    The code uses Gurobi to find a way of assigning salespeople to conventions
    in a way that minimizes the total transportation cost.
    See the sample input and output files for data formats. 
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
    writer=pd.ExcelWriter(outputFile)
    pd.DataFrame([mod.objval],columns=['Minimum cost'])\
        .to_excel(writer,sheet_name='Objective',index=False)
    df=pd.DataFrame(index=I,columns=J)
    for i in I:
        for j in J:
            df.loc[i,j]=x[i,j].x
    df.to_excel(writer,sheet_name='Plan')
    writer.save()
    
if __name__=='__main__':
    import sys, os
    if len(sys.argv)!=3:
        print('Correct syntax: python transportation_optimizer.py inputFile outputFile')
    else:
        inputFile=sys.argv[1]
        outputFile=sys.argv[2]
        if os.path.exists(inputFile):
            transportationOptimizer(inputFile,outputFile)
            print(f'Successfully optimized. Results in "{outputFile}"')
        else:
            print(f'File "{inputFile}" not found!')