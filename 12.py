import pandas as pd

def ding(file):
    
    df = pd.read_excel(file, usecols='D', nrows=23)
    column1 = df.iloc[5:23, 0].tolist()
    df = pd.read_excel(file, usecols='A', nrows=23)
    column2 = df.iloc[5:23, 0].tolist()
    df = pd.read_excel(file, usecols='B', nrows=23)
    column3 = df.iloc[5:23, 0].tolist()
    return column1,column2,column3
    
    
val1,val2,val3 = ding("smvitm.xlsx")

print(val1)
print(val2)
print(val3)