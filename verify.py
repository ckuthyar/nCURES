import pandas as pd
def isempty(file):
    val=0 # determines if the case should be 1 or 0
    df1=pd.read_excel(file,usecols="B",nrows=2)
    c1=df1.iloc[1,0]
    if c1=="" or pd.isna(c1):
        print("Cell is empty")
        val=1
    else:
        print("All good.")
        val=2
    return val

    