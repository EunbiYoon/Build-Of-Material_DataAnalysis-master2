import pandas as pd
import numpy as np

lastweek=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0227/BOM Comparison_TL.xlsx',sheet_name="0220")
thisweek=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0227/BOM Comparison_TL.xlsx',sheet_name="0227")

# empty column fill out
for i in range(len(lastweek)):
    empty_parent=str(lastweek.at[i,"Parent Part"])
    if empty_parent=='nan':
        empty_seq=lastweek.at[i,"Seq."]
        for j in range(len(lastweek)):
            fill_seq=lastweek.at[j,"Seq."]
            fill_level=lastweek.at[j,"level"]
            fill_parent=lastweek.at[j,"Parent Part"]
            fill_part=lastweek.at[j,"Part No"]
            fill_desc=lastweek.at[j,"Desc."]
            if empty_seq==fill_seq:
                lastweek.at[i,"level"]=fill_level
                lastweek.at[i,"Parent Part"]=fill_parent
                lastweek.at[i,"Part No"]=fill_part
                lastweek.at[i,"Desc."]=fill_desc

# empty column fill out
for i in range(len(thisweek)):
    empty_parent=str(thisweek.at[i,"Parent Part"])
    if empty_parent=='nan':
        empty_seq=thisweek.at[i,"Seq."]
        for j in range(len(thisweek)):
            fill_seq=thisweek.at[j,"Seq."]
            fill_level=thisweek.at[j,"level"]
            fill_parent=thisweek.at[j,"Parent Part"]
            fill_part=thisweek.at[j,"Part No"]
            fill_desc=thisweek.at[j,"Desc."]
            if empty_seq==fill_seq:
                thisweek.at[i,"level"]=fill_level
                thisweek.at[i,"Parent Part"]=fill_parent
                thisweek.at[i,"Part No"]=fill_part
                thisweek.at[i,"Desc."]=fill_desc

#file save
lastweek.to_excel('last_fill.xlsx')
thisweek.to_excel('this_fill.xlsx')


# price diff
last_price=lastweek[['match','price match','Parent Part','Part No','Desc.','Parent Item','Child Item','Description']]
this_price=thisweek[['match','price match','Parent Part','Part No','Desc.','Parent Item','Child Item','Description']]
price = pd.merge(this_price,last_price, on=["Parent Item","Child Item"])
for i in range(len(price)):
    this_price=price.at[i,"price match_x"]
    last_price=price.at[i,"price match_y"]
    if this_price!=np.nan and last_price!=np.nan:
        price.at[i,"price_diff"]=round(this_price-last_price,1)

price=price.sort_values(by='price_diff', ascending=False)
price.reset_index(inplace=True, drop=True)
price.to_excel('price_match.xlsx')

# missing part
missing = pd.merge(thisweek, 
                      lastweek, 
                      on ='Part No', 
                      how ='outer')

missing.to_excel('missing_part.xlsx')