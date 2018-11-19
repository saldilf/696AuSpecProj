# 696AuSpecProj
This program reads in absorbance data for gold nanoparticles from an excel file, plots the data (abs vs wavelength), and extracts key parameters from it such as maximum absorbance, wavelength at max absorbance, and size of particle. The program can read from an excel file having multiple columns. 

Everything is automated except the wavelength range. I could make a user command that decides the range, but for my day-to-day purposes and based on how this type of data is reported in the literature, that range works well for gold NPs. For other types of NPs where the nominal spectra will be different from AuNPs, the user would have to modify the code to suit their spectral needs. 

Accompanying the spectra will be a table (or) bar graph of particle size (and other parameters) based on a correlation relating lambdaMax to size and other analyses 

-----------
Things I will add to this program in the future:
  1) user-controlled visualization options. e.g. "do you want subplots or or just one plot, or plot curves X, Y, Z on one plot and A, B, C on another, etc....
  2)Calculate concentration based on user input of extinction coefficinent, vial path length
  3) let the user control what range of wavelengths to plot
