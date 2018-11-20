# 696AuSpecProj
This program does a few things that will save me a lot of time in the future. 

1) Reads in absorbance data for gold nanoparticles from an excel file and plots the data (abs vs wavelength). It can read from multiple columns. 
2) Extracts key parameters from each spectrum such as maximum absorbance, wavelength at max absorbance, and size of particle. 
3) Outputs the plots from (1) and a table summarizing the data in (2)

Everything is automated except the wavelength range. I could make a user command that decides the range, but for my day-to-day purposes and based on how this type of data is reported in the literature, that range works well for gold NPs. For other types of NPs where the nominal spectra will be different from AuNPs, the user would have to modify the code to suit their spectral needs. 

-----------
Things I would like to add to this program in the future:
  1) user-controlled visualization options. e.g. "do you want subplots or just one plot, or plot curves X, Y, Z on one plot and A, B, C on another, etc....
  2) Calculate concentration based on user input of extinction coefficinent, vial path length
  3) let the user control what range of wavelengths to plot
  4) let user decide how legend looks based on how many columns in excel file
  5) user specifies number of sheets in excel file and which ones to analyze
