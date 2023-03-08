import pandas as pd
import numpy as np

lastweek=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0227/BOM Comparison_FL.xlsx',sheet_name="0220")
thisweek=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0227/BOM Comparison_FL.xlsx',sheet_name="0227")

lastweek=lastweek.drop(['Price Change','Substitute Price Change'],axis=1)

result = pd.concat([lastweek, thisweek], axis=1)
result=result.drop(['Unnamed: 0'],axis=1)
price=result[["Parent Item","Child Item","Description",'match','price match']]

price.columns=["Parent Item","Parent Item","Child Item","Child Item","Description","Description","this_match","last_match","this_price","last_price"]


for i in range(len(price)):
    this_price=price.at[i,"this_price"]
    last_price=price.at[i,"last_price"]
    if this_price!=np.nan and last_price!=np.nan:
        price.at[i,"price_diff"]=round(this_price-last_price,1)

price=price.sort_values(by='price_diff', ascending=False)
price.reset_index(inplace=True, drop=True)
price.to_excel('3333.xlsx')