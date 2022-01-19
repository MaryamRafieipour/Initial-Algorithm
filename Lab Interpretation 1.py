import pandas as pd
import math

df_diagnosis = pd.read_excel("CBC.xlsx", sheet_name= "Diagnosis", index_col= "parameters")
df_input = pd.read_excel("input.xlsx", index_col= "parameters")

while True:
  if df_input.at['age','amount'] >= 2 and df_input.at['age','amount'] <= 5:
    table = "2-5YEAR"
    break
  elif df_input.at['age','amount'] > 5 and df_input.at['age','amount'] <= 6:
    table = "5-6YEAR"
    break
  elif df_input.at['age','amount'] > 6 and df_input.at['age','amount'] <= 8:
    table = "6-8YEAR"
    break
  elif df_input.at['age','amount'] > 8 and df_input.at['age','amount'] <= 12:
    table = "8-12YEAR"
    break 
  elif df_input.at['age','amount'] > 12 and df_input.at['age','amount'] <= 18 and df_input.at['gender','amount'] == "female":
    table = "12-18FEMALE"
    break
  elif df_input.at['age','amount'] > 12 and df_input.at['age','amount'] <= 18 and df_input.at['gender','amount'] == "male":
    table = "12-18MALE"
    break
  elif df_input.at['age','amount'] > 18 and df_input.at['gender','amount'] == "female":
    table = "ADULT_FEMALE"
    break  
  elif df_input.at['age','amount'] > 18 and df_input.at['gender','amount'] == "male":
    table = "ADULT_MALE"
    break
  else:
    print("You must enter the gender and age.")
 
df = pd.read_excel("CBC.xlsx", sheet_name= table, index_col= "parameters")

Counter = 0
#HGB---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['HGB','amount']):
  Counter += 1
else:
  HGB = df_input.at['HGB','amount']
  if HGB < df.at['HGB','lower limit']:
    print(df_diagnosis.at['HGB','Diagnosis low'])
  elif HGB > df.at['HGB','upper limit']:
    print(df_diagnosis.at['HGB','Diagnosis up'])
  else:
    Counter += 1
#HCT_PERCENT-------------------------------------------------------------------------------------------
if math.isnan(df_input.at['HCT_PERCENT','amount']):
  Counter += 1
else:
  HCT_PERCENT = df_input.at['HCT_PERCENT','amount']
  if HCT_PERCENT < df.at['HCT_PERCENT','lower limit']:
    print(df_diagnosis.at['HCT_PERCENT','Diagnosis low'])
  elif HCT_PERCENT > df.at['HCT_PERCENT','upper limit']:
    print(df_diagnosis.at['HCT_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#RBC---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['RBC','amount']):
  Counter += 1
else:
  RBC = df_input.at['RBC','amount']
  if RBC < df.at['RBC','lower limit']:
    print(df_diagnosis.at['RBC','Diagnosis low'])
  elif RBC > df.at['RBC','upper limit']:
    print(df_diagnosis.at['RBC','Diagnosis up'])
  else:
    Counter += 1
#MCV---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['MCV','amount']):
  Counter += 1
else:
  MCV = df_input.at['MCV','amount']
  if MCV < df.at['MCV','lower limit']:
    print(df_diagnosis.at['MCV','Diagnosis low'])
  elif MCV > df.at['MCV','upper limit']:
    print(df_diagnosis.at['MCV','Diagnosis up'])
  else:
    Counter += 1
#MCH---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['MCH','amount']):
  Counter += 1
else:
  MCH = df_input.at['MCH','amount']
  if MCH < df.at['MCH','lower limit']:
    print(df_diagnosis.at['MCH','Diagnosis low'])
  elif MCH > df.at['MCH','upper limit']:
    print(df_diagnosis.at['MCH','Diagnosis up'])
  else:
    Counter += 1
#MCHC_PERCENT------------------------------------------------------------------------------------------
if math.isnan(df_input.at['MCHC_PERCENT','amount']):
  Counter += 1
else:
  MCHC_PERCENT = df_input.at['MCHC_PERCENT','amount']
  if MCHC_PERCENT < df.at['MCHC_PERCENT','lower limit']:
    print(df_diagnosis.at['MCHC_PERCENT','Diagnosis low'])
  elif MCHC_PERCENT > df.at['MCHC_PERCENT','upper limit']:
    print(df_diagnosis.at['MCHC_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#RDW_PERCENT-------------------------------------------------------------------------------------------
if math.isnan(df_input.at['RDW_PERCENT','amount']):
  Counter += 1
else:
  RDW_PERCENT = df_input.at['RDW_PERCENT','amount']
  if RDW_PERCENT < df.at['RDW_PERCENT','lower limit']:
    print(df_diagnosis.at['RDW_PERCENT','Diagnosis low'])
  elif RDW_PERCENT > df.at['RDW_PERCENT','upper limit']:
    print(df_diagnosis.at['RDW_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#WBC---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['WBC','amount']):
  Counter += 1
else:
  WBC = df_input.at['WBC','amount']
  if WBC < df.at['WBC','lower limit']:
    print(df_diagnosis.at['WBC','Diagnosis low'])
  elif WBC > df.at['WBC','upper limit']:
    print(df_diagnosis.at['WBC','Diagnosis up'])
  else:
    Counter += 1
#ANC---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['ANC','amount']):
  Counter += 1
else:  
  ANC = df_input.at['ANC','amount']
  if ANC < df.at['ANC','lower limit']:
    print(df_diagnosis.at['ANC','Diagnosis low'])
  elif ANC > df.at['ANC','upper limit']:
    print(df_diagnosis.at['ANC','Diagnosis up'])
  else:
    Counter += 1
#EOS_PERCENT-------------------------------------------------------------------------------------------
if math.isnan(df_input.at['EOS_PERCENT','amount']):
  Counter += 1
else:   
  EOS_PERCENT = df_input.at['EOS_PERCENT','amount']
  if EOS_PERCENT < df.at['EOS_PERCENT','lower limit']:
    print(df_diagnosis.at['EOS_PERCENT','Diagnosis low'])
  elif EOS_PERCENT > df.at['EOS_PERCENT','upper limit']:
    print(df_diagnosis.at['EOS_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#BASOS_PERCENT-----------------------------------------------------------------------------------------
if math.isnan(df_input.at['BASOS_PERCENT','amount']):
  Counter += 1
else: 
  BASOS_PERCENT = df_input.at['BASOS_PERCENT','amount']
  if BASOS_PERCENT < df.at['BASOS_PERCENT','lower limit']:
    print(df_diagnosis.at['BASOS_PERCENT','Diagnosis low'])
  elif BASOS_PERCENT > df.at['BASOS_PERCENT','upper limit']:
    print(df_diagnosis.at['BASOS_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#LYMPHS_PERCENT----------------------------------------------------------------------------------------
if math.isnan(df_input.at['LYMPHS_PERCENT','amount']):
  Counter += 1
else: 
  LYMPHS_PERCENT = df_input.at['LYMPHS_PERCENT','amount']
  if LYMPHS_PERCENT < df.at['LYMPHS_PERCENT','lower limit']:
    print(df_diagnosis.at['LYMPHS_PERCENT','Diagnosis low'])
  elif LYMPHS_PERCENT > df.at['LYMPHS_PERCENT','upper limit']:
    print(df_diagnosis.at['LYMPHS_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#MONOS_PERCENT-----------------------------------------------------------------------------------------
if math.isnan(df_input.at['MONOS_PERCENT','amount']):
  Counter += 1
else: 
  MONOS_PERCENT = df_input.at['MONOS_PERCENT','amount']
  if MONOS_PERCENT < df.at['MONOS_PERCENT','lower limit']:
    print(df_diagnosis.at['MONOS_PERCENT','Diagnosis low'])
  elif MONOS_PERCENT > df.at['MONOS_PERCENT','upper limit']:
    print(df_diagnosis.at['MONOS_PERCENT','Diagnosis up'])
  else:
    Counter += 1
#PLT---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['PLT','amount']):
  Counter += 1
else:   
  PLT = df_input.at['PLT','amount']
  if PLT < df.at['PLT','lower limit']:
    print(df_diagnosis.at['PLT','Diagnosis low'])
  elif PLT > df.at['PLT','upper limit']:
    print(df_diagnosis.at['PLTT','Diagnosis up'])
  else:
    Counter += 1
#MPV---------------------------------------------------------------------------------------------------
if math.isnan(df_input.at['MPV','amount']):
  Counter += 1
else:  
  MPV = df_input.at['MPV','amount']
  if MPV < df.at['MPV','lower limit']:
    print(df_diagnosis.at['MPV','Diagnosis low'])
  elif MPV > df.at['MPV','upper limit']:
    print(df_diagnosis.at['MPV','Diagnosis up'])
  else:
    Counter += 1
#------------------------------------------------------------------------------------------------------
if Counter == 15:
  print("Your blood test is normal.")


