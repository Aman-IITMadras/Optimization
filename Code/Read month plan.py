import pandas as pd
import numpy as np

df1 = pd.read_excel("E:\Aman_doc\EPP Extrusion Plan for the month of Dec'2019.xls")

df1 = df1.rename(columns = {
                                  df1.columns[1]:"priority", 
                                  df1.columns[2]: "container_size",
                                  df1.columns[3]: "material",
                                  df1.columns[5]: "available",
                                  df1.columns[6]: "expected",
                                  }) 


for j in df1.columns:
    x=df1[j].str.contains("PIERCING PRESS")
    for i in range(len(x)):
        if x[i]==True:
            piercing_index=i
            break

df_HE=df1.iloc[:piercing_index,:]
df_HE=df_HE.loc[:,["priority","container_size","material","available","expected"]]

for i in range(len(df_HE["priority"])):
    if isinstance(df_HE["priority"][i],float):
        if np.isnan(df_HE["priority"][i]):
            df_HE=df_HE.drop([i])
    elif isinstance(df_HE["priority"][i],str):
        df_HE=df_HE.drop([i])
 
df_HE = df_HE.reset_index(drop=True)

for i in range(len(df_HE["available"])):
    if df_HE["available"][i]=="-":
        df_HE["available"][i]=0
    elif df_HE["available"][i]==" -":
        df_HE["available"][i]=0
    elif df_HE["available"][i]=="--":
        df_HE["available"][i]=0
    elif df_HE["available"][i]==" --":
        df_HE["available"][i]=0
    elif df_HE["available"][i]=="0":
        df_HE["available"][i]=0
    elif df_HE["available"][i]==" 0":
        df_HE["available"][i]=0
    
for i in range(len(df_HE["expected"])):
    if df_HE["expected"][i]=="-":
        df_HE["expected"][i]=0
    elif df_HE["expected"][i]==" -":
        df_HE["expected"][i]=0
    elif df_HE["expected"][i]=="--":
        df_HE["expected"][i]=0
    elif df_HE["expected"][i]==" --":
        df_HE["expected"][i]=0
    elif df_HE["expected"][i]=="0":
        df_HE["expected"][i]=0
    elif df_HE["expected"][i]==" 0":
        df_HE["expected"][i]=0
        
###################DATA PRE-PROCESSING#####################################
        

