import pandas as pd
import numpy as np

#Read excel
dfm = pd.read_excel("E:\Aman_doc\Productwise production capacity.xlsx")

#Rename columns
dfm = dfm.rename(columns = {dfm.columns[0]:"container_size",
                            dfm.columns[1]:"container",
                                  dfm.columns[2]:"material", 
                                  dfm.columns[3]: "output_size",
                                  dfm.columns[4]: "production",
                                  dfm.columns[5]: "manshifts",
                               }) 

#Separatre Piercing data
piercing_index = dfm[dfm["container_size"].str.contains("PIERCING PRESS")==True].index[0]   
dfm_HE = dfm.iloc[:piercing_index, :]

#Find number of NaN in each column
null_columns=dfm_HE.columns[dfm_HE.isnull().any()]
x=dfm_HE[null_columns].isnull().sum()

#If "number of zeros in the column" is equal ro number of rows, i.e whole column is empty
for i in range(len(x)):
    if x[i]==len(dfm_HE.index):
        first_zero_column = i
        break

#Remove columns after the empty column  
dfm_HE = dfm_HE.iloc[:,:first_zero_column]
        
#Find for first integer or float number(except NaN) in "production" column, and take the row number
for i in range(len(dfm_HE["production"])):
    if isinstance(dfm_HE["production"][i], int):
        start = i
        break
    elif isinstance(dfm_HE["production"][i], float):
        if np.invert(np.isnan(dfm_HE["production"][i])) :
            start = i
            break


for i in range(len(dfm_HE["production"])):
    if isinstance(dfm_HE["production"][len(dfm_HE["production"])-i-1], int):
        end = len(dfm_HE["production"])-i-1
        break
    elif isinstance(dfm_HE["production"][len(dfm_HE["production"])-i-1], float):
        if np.invert(np.isnan(dfm_HE["production"][len(dfm_HE["production"])-i-1])) :
            end = len(dfm_HE["production"])-i-1
            break
        
dfm_HE = dfm_HE.iloc[start:end+1,:]
dfm_HE = dfm_HE.reset_index(drop=True)

dfm_HE = dfm_HE.loc[:,["container_size", "container","material","output_size","production","manshifts"]]

#makes output_size column values "True" if it has "x" or "X" i.e it is a pipe, an nan if it is a rod
a=dfm_HE["output_size"].str.contains("x")
b=dfm_HE["output_size"].str.contains("X")

for i in range(len(dfm_HE["output_size"])):
    if a[i]==True:
        dfm_HE["output_size"][i]=True
    elif b[i]==True:
        dfm_HE["output_size"][i]=True
    else:
         dfm_HE["output_size"][i]=False
        
###################DATA PRE-PROCESSING#####################################


    