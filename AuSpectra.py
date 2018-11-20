#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 15:00:19 2018

@author: salwanbutrus


"""
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import pandas as pd
import collections as cl
import six



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
            axes = plt.gca()
            box = axes.get_position()
            axes.set_position([box.x0, box.y0, box.width * 0.983, box.height])
            plt.xlabel('Wavelength (nm)')
            plt.ylabel('Absorbance (Normalized to Amax)')
        
        plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
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
            axes = plt.gca()
            box = axes.get_position()
            axes.set_position([box.x0, box.y0, box.width * 0.983, box.height]) #to fit legend in output
            plt.xlabel('Wavelength (nm)')
            plt.ylabel('Absorbance')
            
    
        plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
        axes = plt.gca()
        box = axes.get_position()
        axes.set_position([box.x0, box.y0, box.width * 0.983, box.height])
        axes.set_xlim([400 , 700])
        axes.set_ylim([0 , Amax + 0.05])
   
        
    else:
    	print("Please enter 'Yes' or 'No' ")

#format data dict into a frame
df = pd.DataFrame(data)  


#output data frame as nice table
def render_mpl_table(data, col_width=4.0, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')

    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)

    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in  six.iteritems(mpl_table._cells):
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    return ax

render_mpl_table(df, header_columns=0, col_width=3.0)

plt.show() #show all plots before this line

