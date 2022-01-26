import pandas as pd
import math

df_diagnosis = pd.read_excel("CBC_Adult.xlsx", sheet_name= "Diagnosis", index_col= "parameters")
df_input = pd.read_excel("SAMPLE.xlsx", index_col= "parameters")


if df_input.at['age','amount'] < 18:
  print("Age of test taker must be over 18 years old")


if df_input.at['gender','amount'] == "female":
    df = pd.read_excel("CBC_Adult.xlsx", sheet_name= "Female", index_col= "parameters")
elif df_input.at['age','amount'] > 12 and df_input.at['age','amount'] <= 18 and df_input.at['gender','amount'] == "male":
    df = pd.read_excel("CBC_Adult.xlsx", sheet_name= "Male", index_col= "parameters")

Counter = 0
#Polycythemia------------------------------------------------------------------------------------------
if math.isnan(df_input.at['HGB','amount'] and df_input.at['HCT_PERCENT','amount']):
  Counter += 1
else:
  HGB = df_input.at['HGB','amount']
  HCT_PERCENT = df_input.at['HCT_PERCENT','amount']
  if HGB > df.at['HGB','upper limit'] and HCT_PERCENT > df.at['HCT_PERCENT','upper limit']:
    print(df_diagnosis.at['Polycythemia','Diagnosis'])
  else:
    Counter += 1
#Eosinophilia------------------------------------------------------------------------------------------
if math.isnan(df_input.at['AEC','amount']):
  Counter += 1
else:
  AEC = df_input.at['AEC','amount']
  if AEC > df.at['AEC','upper limit']:
    print(df_diagnosis.at['Eosinophilia','Diagnosis'])
  else:
    Counter += 1
#Lymphocytosis-----------------------------------------------------------------------------------------
if math.isnan(df_input.at['ALC','amount']):
  Counter += 1
else:
  ALC = df_input.at['ALC','amount']
  if ALC < df.at['ALC','upper limit']:
    print(df_diagnosis.at['Lymphocytosis','Diagnosis'])
    Counter += 1
#Lymphocytopenia---------------------------------------------------------------------------------------
if math.isnan(df_input.at['ALC','amount']):
  Counter += 1
else:
  ALC = df_input.at['ALC','amount']
  if ALC < df.at['ALC','lower limit']:
    print(df_diagnosis.at['Lymphocytopenia','Diagnosis'])
  else:
    Counter += 1
#Neutropenia-------------------------------------------------------------------------------------------
if math.isnan(df_input.at['ANC','amount']):
  Counter += 1
else:
  ANC = df_input.at['ANC','amount']
  if ANC < df.at['ANC','lower limit']:
    print(df_diagnosis.at['Neutropenia','Diagnosis'])
  else:
    Counter += 1
#Thrombocytosis----------------------------------------------------------------------------------------
if math.isnan(df_input.at['PLT','amount']):
  Counter += 1
else:
  PLT = df_input.at['PLT','amount']
  if PLT > df.at['PLT','upper limit']:
    print(df_diagnosis.at['Thrombocytosis','Diagnosis'])
  else:
    Counter += 1   
#Thrombocytopenia -------------------------------------------------------------------------------------
if math.isnan(df_input.at['PLT','amount']):
  Counter += 1
else:
  PLT = df_input.at['PLT','amount']
  if PLT < df.at['PLT','lower limit']:
    print(df_diagnosis.at['Thrombocytopenia','Diagnosis'])
  else:
    Counter += 1    
#------------------------------------------------------------------------------------------------------
if Counter == 7:
  print("Your blood test is normal.")


