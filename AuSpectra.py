#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 15:00:19 2018

@author: salwanbutrus


Utilities
    1) Size range restriction 

Testing: 
    1) opening file
        a) not xlsx
        b) not in folder we're in now
    2)


"""

import numpy as np
import matplotlib.pyplot as plt
import xlrd
import pandas as pd
import collections as cl


#prompt user to open file
path = input("Enter a file name (e.g. data.xlsx): ")
wb = xlrd.open_workbook(path)
sheet1 = wb.sheet_by_index(0)

#dimensions of file
r = sheet1.nrows
c = sheet1.ncols

#wavelength range of data-
xLimLflt = sheet1.cell(1,0).value #must be less than 400
xLimL = int(xLimLflt)

xLimUflt = sheet1.cell(r-1,0).value
xLimU = int(xLimUflt)

#dict for extracted or calculated values from data
data = cl.OrderedDict({'Sample ID': [],
                       'lambdaMax (nm)': [], 
                       'Amax': [],
                       'Size (nm)': [],
                       'Concentration': []
                       })

#for loop that plots/analyzes based on how many columns     
norm = None
while norm not in ("Yes", "No"):
    
    norm = input("Would you like to normalize the data to the maximum absorbance? (Enter 'Yes' or 'No') ")
    
    if norm == "Yes":
        for x in range(3,c):
            #299 is the number of data points minus 2
            lambdas = sheet1.col_values(colx = 0,start_rowx = (xLimU - 299 ) - xLimL,  end_rowx = r) #x-vals
            abso = sheet1.col_values(colx = x,start_rowx = (xLimU - 299)-xLimL,end_rowx = r) #y-vals loop cycles thru
            
            #400 is the lambda we start plotting at
            sID = sheet1.cell(0,x).value
            lMax = abso.index(max(abso)) + 400
            Amax = max(abso)
           
            absoNorm = [x/Amax for x in abso] #normalizes to Amax
            
            size = -0.02111514*(lMax**2.0) + 24.6*(lMax) - 7065.
            #J. Phys. Chem. C 2007, 111, 14664-14669
            
            if 518 < lMax < 570:
                data['Size (nm)'].append(int(size))
    
            else:
                data['Size (nm)'].append('>100')

            
            data['Sample ID'].append(sID)
            data['lambdaMax (nm)'].append(lMax)
            data['Amax'].append( Amax )
            data['Concentration'].append('placeholder')
        
            plt.plot(lambdas, absoNorm ,linewidth=2,label= sheet1.cell(0,x).value)
            plt.xlabel(sheet1.cell(0,0).value)
            plt.ylabel('Absorbance (Normalized to Amax)')
        
       # plt.legend()
        plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
        axes = plt.gca()
        axes.set_xlim([400 , 700])
        axes.set_ylim([0 , 1.5])
                
        
        
    elif norm == "No":
        
        for x in range(3,c):
            #299 is the number of data points minus 2
            lambdas = sheet1.col_values(colx = 0,start_rowx = (xLimU - 299 ) - xLimL,  end_rowx = r) #x-vals
            abso = sheet1.col_values(colx = x,start_rowx = (xLimU - 299)-xLimL,end_rowx = r) #y-vals loop cycles thru
            
            #400 is the lambda we start plotting at
            sID = sheet1.cell(0,x).value
            lMax = abso.index(max(abso)) + 400
            Amax = max(abso)
            
            size = -0.02111514*(lMax**2.0) + 24.6*(lMax) - 7065.
            #J. Phys. Chem. C 2007, 111, 14664-14669
            
            if 518 < lMax < 570:
                data['Size (nm)'].append(int(size))
    
            else:
                data['Size (nm)'].append('>100')
            
            data['Sample ID'].append(sID)
            data['lambdaMax (nm)'].append(lMax)
            data['Amax'].append( Amax )
            data['Concentration'].append('placeholder')
        
            plt.plot(lambdas, abso ,linewidth=2,label= sheet1.cell(0,x).value)
            plt.xlabel(sheet1.cell(0,0).value)
            plt.ylabel('Absorbance')
            
    
        plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
        axes = plt.gca()
        axes.set_xlim([400 , 700])
        axes.set_ylim([0 , Amax + 0.05])
           
        
    else:
    	print("Please enter yes or no.")
    
    
frame = pd.DataFrame(data)
print(frame)







#plt.savefig('testfig.png', format='png', dpi=1000)







    





