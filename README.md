# fatima_alhourani_et_al
Drug synergy calculation



# Code information

## Code author

Diego Tosi


Code Version 2022
R version > 4.1.2

Dependancies: 

readxl,
gplots,
dplyr,
gtools,
gridExtra,
gridGraphics,
ggplot2,
ggpubr,
plot3D,
reshape2,
openxlsx,

## Data file format specifications

You must provide a xlsx file with one or more sheets

The sheets names must be different from each other

In each sheet:
 1. The first row must contains the names of two drugs in the first two cells
 2. Drug name order: row drug, column drug
 3. The second row must be blank
 4. If two or more data matrices : always separated by a blank row

 The name of figure will be composed by:
   1. the name of the corresponding sheet
   2. the rank of the corresponding matrix in the sheet

 Create a folder named "output" in the wd, in order to collect the figure files
 Modify the settings in the corresponding section here below and it's done

