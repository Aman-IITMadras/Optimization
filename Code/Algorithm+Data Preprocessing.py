import pandas as pd
import numpy as np
import os

c_s_n = int(input("Which size contianer is IN USE:\n1) 150\n2) 180\n3) 225\n4) 300\n5) 350\nEnter '1' if 150 container is IN USE\nEnter '2' if 180 container is IN USE...\n"))
if c_s_n == 1:
    container_attached = 150
elif c_s_n == 2:
    container_attached = 180
elif c_s_n == 3:
    container_attached = 225
elif c_s_n == 4:
    container_attached = 300   
elif c_s_n == 5:
    container_attached = 350   
else:
    print("\nNO, such container")

c_n = int(input("\nWhich Container is IN USE:\n1) SS\n2) Zr\nEnter '1' if SS\nEnter '2' if Zr\n"))
if c_n==1:
    container_material = "SS"
elif c_n==2:
    container_material = "Zr"
else:
    print("\nNO, such container")

max_350_Zr_148_152_slab=[0,0,0]
max_150_Zr_rod_tube=[0,0]

#Input
max_150_Zr_rod_tube[0]=int(input("\nHow many (maximum) to be extruded\nin 150, Zr container RODS\n"))
max_150_Zr_rod_tube[1]=int(input("\nHow many (maximum) to be extruded\nin 150, Zr container TUBES\n"))

max_350_Zr_148_152_slab[0]=int(input("\nHow many (maximum) to be extruded\nin 350, Zr container 148mm\n"))
max_350_Zr_148_152_slab[1]=int(input("\nHow many (maximum) to be extruded\nin 350, Zr container 152mm\n"))
max_350_Zr_148_152_slab[2]=int(input("\nHow many (maximum) to be extruded\nin 350, Zr container SLAB\n"))

  
#Take Input
#container_attached = 300
#container_material = "SS"

for i in os.listdir("C:/Users/arora/Desktop/Production Plan©/[UPLOAD] NFC Calendar"):
    if (".ini" not in i):
        path_calendar="C:/Users/arora/Desktop/Production Plan©/[UPLOAD] NFC Calendar/"+i

for i in os.listdir("C:/Users/arora/Desktop/Production Plan©/[UPLOAD] Available Material Stock & Product requirements"):
    if (".ini" not in i):
        path_month_plan="C:/Users/arora/Desktop/Production Plan©/[UPLOAD] Available Material Stock & Product requirements/"+i
    
for i in os.listdir("C:/Users/arora/Desktop/Production Plan©/[UPLOAD] Production Rate & Manshift Data"):
    if (".ini" not in i):
        path_production_capacity="C:/Users/arora/Desktop/Production Plan©/[UPLOAD] Production Rate & Manshift Data/"+i

df1 = pd.read_excel(path_month_plan)

df1 = df1.rename(columns = {
                                  df1.columns[1]:"priority", 
                                  df1.columns[2]: "container_size",
                                  df1.columns[3]: "material",
                                  df1.columns[4]: "dimension",
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
df_HE=df_HE.loc[:,["priority","container_size","material","dimension","available","expected"]]

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
        
a=(df_HE["dimension"].str.contains("x"))|(df_HE["dimension"].str.contains("X"))

df_HE["tube"]=a        
###################END of DATA PRE-PROCESSING of Priority month plan#####################################

#Read excel
dfm = pd.read_excel(path_production_capacity)

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
        
###################END of DATA PRE-PROCESSING of Material Time#####################################

#Add column of "Zr", "SS" container to priority list
error=0        
container=[]
for index,i in enumerate(df_HE["material"]):
    if error==index:
        for index_,j in enumerate(dfm_HE["material"]):
            if i==j:
                container.append(dfm_HE["container"][index_])
                error=error+1
                break
    else:
         print("Error!!! Material not found in the list")
         print("Please verify if the material you have used are among:")
         print("Zr-4, Zr-2, ZNC, Zr1%Nb, PT7M, SS-304L, SS-316L, SS-321, D-9, SuperNi, STA-59") 
         print("9Cr1Mb, Inc-800, ZrNb, SS-304L, SS-321, Inc-617, Ti Half Alloy, SS-304L, MDN-350, MDN-250")
         print('Make sure you have not MISSESD ANY "-" OR Left extra space in material you have mentioned OR interchanged lower and upper case letters')
         break
                
#df_HE["container"] = container

df_HE.insert(2,"container",container)

#Read calendar
df_calendar=pd.read_excel(path_calendar)

#Find number of working days
df_working=df_calendar[df_calendar["working day"]==1]["Date"]
df_working=df_working.reset_index(drop=True)
working_days=len(df_working)
for index, i in enumerate(df_working):
    j=str(i)
    df_working.update(pd.Series([j[:11]], index = [index]))

#Add columns to df_HE, "days_required"=>Number of days the material will take
                      #"production"=> Per day production of the material
error=0 
days_required=[]
daily_production=[]
manshift=[]
total_production=0
total_manshift=0
for i in range(len(df_HE)):
    if error==i:
        for j in range(len(dfm_HE)):
            if (df_HE["material"][i]==dfm_HE["material"][j])&(df_HE["container_size"][i]==dfm_HE["container_size"][j])&(df_HE["tube"][i]==dfm_HE["output_size"][j]):
                days_required.append(float(df_HE["available"][i]/dfm_HE["production"][j]))
                daily_production.append(float(dfm_HE["production"][j]))
                manshift.append((float(dfm_HE["manshifts"][j])*dfm_HE["production"][j])/100.0)
                error=error+1
                break
    else:
         print("Error!!! Material not found in the list")
         print("Please verify if the material you have used are among:")
         print("Zr-4, Zr-2, ZNC, Zr1%Nb, PT7M, SS-304L, SS-316L, SS-321, D-9, SuperNi, STA-59") 
         print("9Cr1Mb, Inc-800, ZrNb, SS-304L, SS-321, Inc-617, Ti Half Alloy, SS-304L, MDN-350, MDN-250")
         print('Make sure you have not MISSESD ANY "-" OR Left extra space in material you have mentioned OR interchanged lower and upper case letters')
         break

if len(days_required)==len(df_HE):
    df_HE["days required"]=days_required
if len(days_required)==len(df_HE):
    df_HE["daily production"]=daily_production
if len(days_required)==len(df_HE):
    df_HE["manshift"]=manshift

#From df_HE taking rows of attached container    
#df_attached_container=df_HE[(df_HE["container_size"]==container_attached)&(df_HE["container"]==container_material)&((df_HE["priority"]==1)|(df_HE["priority"]==2))]

#Defining varible days_remaining which indicates days remain in month after each operation
days_remaining=working_days

#Defining "plan" which will be the plan output after appending rows 
plan = pd.DataFrame(columns=["container_size","container","material","dimension","expected production","manshift"])

#Adding rows into "plan" for attached conatiner
#for index,i in enumerate(df_attached_container["days required"]):
    #for j in range(int(i)):
       # plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index]}, ignore_index=True)
       # days_remaining=days_remaining-1

#container already used
df_containers_used=set()
    
df_container_priority=["150Zr","350Zr","150SS","180Zr","180SS","225Zr","225SS","300Zr","300SS","350SS"]

def containerNumber(c_s,c):
    y=str(c_s)+c
    for index,i in enumerate(df_container_priority):
        if i==y:
            return index      

def changeContainer(c_s,c):
    df_containers_used.add(containerNumber(c_s,c))
    if ({0,1,2,3,4,5,6,7,8,9}-df_containers_used)==set():
        return None
    else:
        return min({0,1,2,3,4,5,6,7,8,9}-df_containers_used)

while days_remaining>1:
    atleastOne=False
    #Adding rows into "plan" for attached conatiner
    df_attached_container=df_HE[(df_HE["container_size"]==container_attached)&(df_HE["container"]==container_material)]
    df_attached_container=df_attached_container.reset_index(drop=True)
    container_used=0 #Number of time the current SS container is used
    container_used_rod_slab_tube=[0,0,0]
    for index,i in enumerate(df_attached_container["days required"]):
        if days_remaining<1:
            break
        else:
            for j in range(int(i)):
                if days_remaining<1:
                    break
                else:
                    if container_material=="SS":
                        container_used=container_used+df_attached_container["daily production"][index]
                        if container_used<800:
                            plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                            atleastOne=True
                            total_production = total_production + df_attached_container["daily production"][index]
                            total_manshift = total_manshift + df_attached_container["manshift"][index]
                            days_remaining=days_remaining-1
                        else:
                            break
                    elif (container_material=="Zr")&(container_attached==350):
                        if df_attached_container["tube"][index]==True:
                            container_used_rod_slab_tube[1]=container_used_rod_slab_tube[1]+df_attached_container["daily production"][index]
                            if container_used_rod_slab_tube[1]<max_350_Zr_148_152_slab[2]:
                                plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                                atleastOne=True
                                total_production = total_production + df_attached_container["daily production"][index]
                                total_manshift = total_manshift + df_attached_container["manshift"][index]
                                days_remaining=days_remaining-1
                            else:
                                break
                        elif "148" in df_attached_container["dimension"][index]:
                            container_used_rod_slab_tube[2]=container_used_rod_slab_tube[2]+df_attached_container["daily production"][index]
                            if container_used_rod_slab_tube[2]<max_350_Zr_148_152_slab[0]:
                                plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                                atleastOne=True
                                total_production = total_production + df_attached_container["daily production"][index]
                                total_manshift = total_manshift + df_attached_container["manshift"][index]
                                days_remaining=days_remaining-1
                            else:
                                break
                        elif "152" in df_attached_container["dimension"][index]:
                            container_used_rod_slab_tube[0]=container_used_rod_slab_tube[0]+df_attached_container["daily production"][index]
                            if container_used_rod_slab_tube[0]<max_350_Zr_148_152_slab[1]:
                                plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                                atleastOne=True
                                total_production = total_production + df_attached_container["daily production"][index]
                                total_manshift = total_manshift + df_attached_container["manshift"][index]
                                days_remaining=days_remaining-1
                            else:
                                break
                    elif (container_material=="Zr")&(container_attached==150):
                        if df_attached_container["tube"][index]==True:
                            container_used_rod_slab_tube[2]=container_used_rod_slab_tube[2]+df_attached_container["daily production"][index]
                            if container_used_rod_slab_tube[2]<max_150_Zr_rod_tube[1]:
                                plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                                atleastOne=True
                                total_production = total_production + df_attached_container["daily production"][index]
                                total_manshift = total_manshift + df_attached_container["manshift"][index]
                                days_remaining=days_remaining-1
                            else:
                                break
                        elif df_attached_container["tube"][index]==False:
                            container_used_rod_slab_tube[0]=container_used_rod_slab_tube[0]+df_attached_container["daily production"][index]
                            if container_used_rod_slab_tube[0]<max_150_Zr_rod_tube[0]:
                                plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift': round(df_attached_container["manshift"][index])}, ignore_index=True)
                                atleastOne=True
                                total_production = total_production + df_attached_container["daily production"][index]
                                total_manshift = total_manshift + df_attached_container["manshift"][index]
                                days_remaining=days_remaining-1
                            else:
                                break
                    else:
                         plan = plan.append({'container_size': df_attached_container["container_size"][index], 'container': df_attached_container["container"][index], 'material': df_attached_container["material"][index], 'dimension': df_attached_container["dimension"][index], 'expected production': df_attached_container["daily production"][index], 'manshift':round(df_attached_container["manshift"][index])}, ignore_index=True)
                         atleastOne=True
                         total_production = total_production + df_attached_container["daily production"][index]
                         total_manshift = total_manshift + df_attached_container["manshift"][index]
                         days_remaining=days_remaining-1
                        
    #decides whcih container to use next    
    if changeContainer(container_attached,container_material)==None:
        break
    next_container=df_container_priority[changeContainer(container_attached,container_material)]
    container_attached=int(next_container[0:3])
    container_material=str(next_container[3:])
    if days_remaining >1:
        tool="Tool Change to "+next_container
        if atleastOne==True:
            plan = plan.append({'container_size': tool}, ignore_index=True)
            days_remaining=days_remaining-1
        else:
            plan.drop(plan.tail(1).index,inplace=True)
            plan = plan.append({'container_size': tool}, ignore_index=True)
            days_remaining=days_remaining-1
            
plan=pd.concat([df_working,plan], axis=1)
plan.loc[len(plan)]= [None,None,None,None,"Total",round(total_production),round(total_manshift)]



output_path = "C:/Users/arora/Desktop/Production Plan©/[OUTPUT] Production Plan/Month Plan.xlsx"
plan.to_excel(output_path) 
print("You will find the Production Plan Excel at:\n",output_path)
print("Total Production:", round(total_production))
print("Total Manshift:", round(total_manshift)) 
input()
    
   
        