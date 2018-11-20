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


#prompt user to open file and catch erros  
wb = None
    
while not wb:
    try:
        wb = xlrd.open_workbook(input("Enter a file name (e.g. data.xlsx):  "))
        
    except FileNotFoundError:
        print("File not found in this directory. Make sure it is in the directory and follows 'data.xlsx' ")
        wb = None


sheet1 = wb.sheet_by_index(0) #define sheet you will analyze

#dimensions of sheet
r = sheet1.nrows
c = sheet1.ncols

#wavelength range of data
#Lower limit
xLimLflt = sheet1.cell(1,0).value #must be less than 400
xLimL = int(xLimLflt)

#Upper limit
xLimUflt = sheet1.cell(r-1,0).value
xLimU = int(xLimUflt)

#dict for extracted or calculated values from data
data = cl.OrderedDict({'Sample ID': [],
                       'lambdaMax (nm)': [], 
                       'Amax': [],
                       'Size (nm)': []
                       })

#for loop that plots/analyzes based on how many columns     
norm = None
while norm not in ("Yes", "No"):
    
    norm = input("Would you like to normalize the data to the maximum absorbance? (Enter 'Yes' or 'No') ")
    
    if norm == "Yes":
        for x in range(3,c):
            #299 is the number of data points minus 2
            lambdas = sheet1.col_values(colx = 0,start_rowx = (xLimU - 299 ) - xLimL,  end_rowx = r) #x-vals (wavelength)
            abso = sheet1.col_values(colx = x,start_rowx = (xLimU - 299)-xLimL,end_rowx = r) #y-vals (absorbance)
            
            #400 is the lambda we start plotting at
            sID = sheet1.cell(0,x).value #name of each column
            lMax = abso.index(max(abso)) + 400 #find max lambda for a column
            Amax = max(abso) #get max Abs to norm against it
           
            absoNorm = [x/Amax for x in abso] #normalizes to Amax
            
            size = -0.02111514*(lMax**2.0) + 24.6*(lMax) - 7065.
            #J. Phys. Chem. C 2007, 111, 14664-14669
            
            #if lambdaMax outisde of this range then 1) particles aggregated and 2)outside of correlation
            if 518 < lMax < 570:
                data['Size (nm)'].append(int(size))
    
            else:
                data['Size (nm)'].append('>100')

            #added extracted/calculated items into dict 'data'
            data['Sample ID'].append(sID)
            data['lambdaMax (nm)'].append(lMax)
            data['Amax'].append( Amax )
        
            #plot each column (cycle in loop)
            plt.plot(lambdas, absoNorm ,linewidth=2,label= sheet1.cell(0,x).value)
            plt.xlabel(sheet1.cell(0,0).value)
            plt.ylabel('Absorbance (Normalized to Amax)')
        
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
            sID = sheet1.cell(0,x).value #name of each column
            lMax = abso.index(max(abso)) + 400 #find max lambda for a column
            Amax = max(abso) #get max Abs to set y-lim

            size = -0.02111514*(lMax**2.0) + 24.6*(lMax) - 7065.
            #J. Phys. Chem. C 2007, 111, 14664-14669
            
            
            #if lambdaMax outisde of this range then 1) particles aggregated and 2)outside of correlation
            if 518 < lMax < 570:
                data['Size (nm)'].append(int(size))
    
            else:
                data['Size (nm)'].append('>100')
            
            #added extracted/calculated items into dict 'data'
            data['Sample ID'].append(sID)
            data['lambdaMax (nm)'].append(lMax)
            data['Amax'].append( Amax )
        
            plt.plot(lambdas, abso ,linewidth=2,label= sheet1.cell(0,x).value)
            plt.xlabel(sheet1.cell(0,0).value)
            plt.ylabel('Absorbance')
            
    
        plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
        axes = plt.gca()
        axes.set_xlim([400 , 700])
        axes.set_ylim([0 , Amax + 0.05])
        plt.show() #show plot from command line
   
        
    else:
    	print("Please enter 'Yes' or 'No' ")
    
    
#format data dict into a frame
frame = pd.DataFrame(data)
print(frame)
print("MAKE INTO A NICE TABLE")

IDs = data['Sample ID']
sizes= data['Size (nm)']

print(len(IDs))

#plt.bar(IDs, sizes)
#plt.legend()



#print(IDs[0])
#print(IDs, sizes)






#plt.savefig('testfig.png', format='png', dpi=1000)







    





