#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov 17 15:00:19 2018

@author: salwanbutrus


At a minimum, there should be two options that the user can choose to control the output of the program. 
    1) No plot: type in file name to get info on spectrum (e.g. "sample A aggregated or is 58 nm based on correlation."
     can get the info as a table output. 
         a) normalized to highest value
         b) not normalized 
    2) Do #1 above AND plot all the spectra and label them. 
The info in #1 can just be attached to each spectrum rather than in a table, or both. 

TO DO:
1)put spectra in different csv files and label them. Put a test spectrum for now 



"""

import numpy as np
from scipy.integrate import odeint
import matplotlib.pyplot as plt
import xlrd

path = ('/Users/salwanbutrus/Desktop/Desktop_Organized/ChE696/Project/TemplateData.xlsx')
wb = xlrd.open_workbook(path)
sheet1 = wb.sheet_by_index(0)

r = sheet1.nrows
c = sheet1.ncols

print(r,c)

value = sheet1.cell(0, 0)

print(value)

lambdas = sheet1.col_values(colx = 0,start_rowx = 1,end_rowx = r)

#now  for loop the plots based on how many columns


#plot label should be the name of the data in column


#Take user input for file
