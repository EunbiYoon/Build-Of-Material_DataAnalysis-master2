import pandas as pd
import numpy as np

###########################GERP###############################
gerp=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/gerp.xlsx')
# #파일에 따라 추가하기
# gerp=gerp[gerp['Model'].str.contains('F3')] #total 행 제거
# gerp.reset_index(drop=True, inplace=True)
for i in range(len(gerp)):
    level=gerp.at[i,'Level']
    if level==0:
        gerp.at[i,'level0']=gerp.at[i,'Child Item']
    elif level=='*1':
        gerp.at[i,'level0']=gerp.at[i,'Parent Item']
        gerp.at[i,'level1']=gerp.at[i,'Child Item']
    elif level=='**2':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level1']=gerp.at[i,'Parent Item']
        gerp.at[i,'level2']=gerp.at[i,'Child Item']
    elif level=='***3':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level2']=gerp.at[i,'Parent Item']
        gerp.at[i,'level3']=gerp.at[i,'Child Item']
    elif level=='****4':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level3']=gerp.at[i,'Parent Item']
        gerp.at[i,'level4']=gerp.at[i,'Child Item']
    elif level=='*****5':
        gerp.at[i,'level0']=gerp.at[0,'Parent Item'] #가장 첫번째
        gerp.at[i,'level4']=gerp.at[i,'Parent Item']
        gerp.at[i,'level5']=gerp.at[i,'Child Item']
    else:
        pass

#Level 3 - level1
empty_list=gerp.loc[gerp['Level']=='***3'].reset_index()
full_list=gerp.loc[gerp['Level']=='**2'].reset_index()
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level2']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level2']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#Level 3 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value=empty_list.at[i,"level1"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value
        else:
            pass
#Level 4 - level2, level1
empty_list=gerp.loc[gerp['Level']=='****4'].reset_index()
full_list=gerp.loc[gerp['Level']=='***3'].reset_index()   
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level3']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level3']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#Level 4 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value1
            gerp.at[j,"level2"]=empty_value2
        else:
            pass
#Level 5 - level3,level2, level1
empty_list=gerp.loc[gerp['Level']=='*****5'].reset_index()
full_list=gerp.loc[gerp['Level']=='****4'].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level4']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level4']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']

#Level 5 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    for j in range(len(gerp)):
        if empty_index==j:
            gerp.at[j,"level1"]=empty_value1
            gerp.at[j,"level2"]=empty_value2
            gerp.at[j,"level3"]=empty_value3
        else:
            pass

gerp.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/gerpresult.xlsx')

###########################NPT###############################
npt=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/npt.xlsx')

# #파일에 따라 추가하기 (Test-1 O, Test-3 X)
npt=npt.drop([0],axis=0) #total 행 제거
npt.reset_index(drop=True, inplace=True)

#NPT level 숫자로 정리
for i in range(len(npt)):
    #level 정리
    level=str(npt.at[i,"Lvl"])
    level=int(level.replace('.',''))
    npt.at[i,"level"]=level

#parent part No 생성하기
for i in range(len(npt)):
    level=npt.at[i,"level"]
    #level로 순위 메기기
    if level!=0:
        for j in range(len(npt)):
            compare_level=npt.at[i-j,"level"]
            if compare_level<level:
                Parent_Part=npt.at[i-j,"Part No"]
                break
            else:
                continue
        npt.at[i,"Parent Part"]=Parent_Part
    else:
        #부모는 자기자신
        npt.at[i,"Parent Part"]=npt.at[i,"Part No"]

#level0,level1,level2,level3, level4, level5 생성하기
for i in range(len(npt)):
    level=npt.at[i,'level']
    if level==0:
        npt.at[i,'level0']=npt.at[i,'Part No']
    elif level==1:
        npt.at[i,'level0']=npt.at[i,'Parent Part']
        npt.at[i,'level1']=npt.at[i,'Part No']
    elif level==2:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level1']=npt.at[i,'Parent Part']
        npt.at[i,'level2']=npt.at[i,'Part No']
    elif level==3:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level2']=npt.at[i,'Parent Part']
        npt.at[i,'level3']=npt.at[i,'Part No']
    elif level==4:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level3']=npt.at[i,'Parent Part']
        npt.at[i,'level4']=npt.at[i,'Part No']
    elif level==5:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level4']=npt.at[i,'Parent Part']
        npt.at[i,'level5']=npt.at[i,'Part No']
    elif level==6:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level5']=npt.at[i,'Parent Part']
        npt.at[i,'level6']=npt.at[i,'Part No']
    elif level==7:
        npt.at[i,'level0']=npt.at[0,'Parent Part'] #가장 첫번째
        npt.at[i,'level6']=npt.at[i,'Parent Part']
        npt.at[i,'level7']=npt.at[i,'Part No']
    else:
        pass

#Level 3 - level1
empty_list=npt.loc[npt['level']==3].reset_index()
full_list=npt.loc[npt['level']==2].reset_index()
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level2']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level2']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 3 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value=empty_list.at[i,"level1"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value
        else:
            pass

#level 4 - level2, level1
empty_list=npt.loc[npt['level']==4].reset_index()
full_list=npt.loc[npt['level']==3].reset_index()   
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level3']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level3']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 4 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
        else:
            pass

#level 5 - level3,level2, level1
empty_list=npt.loc[npt['level']==5].reset_index()
full_list=npt.loc[npt['level']==4].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level4']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level4']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 5 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
        else:
            pass

#level 6 - level5,level4,level3,level2,level1
empty_list=npt.loc[npt['level']==6].reset_index()
full_list=npt.loc[npt['level']==5].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level5']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level5']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level4"]=full_list.at[j,'level4']
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 6 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    empty_value4=empty_list.at[i,"level4"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
            npt.at[j,"level4"]=empty_value4
        else:
            pass

#level 7 - level6,level5,level4,level3,level2,level1
empty_list=npt.loc[npt['level']==7].reset_index()
full_list=npt.loc[npt['level']==6].reset_index() 
for i in range(len(empty_list)):
    empty_value=empty_list.at[i,'level6']
    for j in range(len(full_list)):
        full_value=full_list.at[j,'level6']
        if empty_value==full_value:
            break
        else:
            continue
    empty_list.at[i,"level5"]=full_list.at[j,'level5']
    empty_list.at[i,"level4"]=full_list.at[j,'level4']
    empty_list.at[i,"level3"]=full_list.at[j,'level3']
    empty_list.at[i,"level2"]=full_list.at[j,'level2']
    empty_list.at[i,"level1"]=full_list.at[j,'level1']
#level 7 - put back to original
for i in range(len(empty_list)):
    empty_index=empty_list.at[i,"index"]
    empty_value1=empty_list.at[i,"level1"]
    empty_value2=empty_list.at[i,"level2"]
    empty_value3=empty_list.at[i,"level3"]
    empty_value4=empty_list.at[i,"level4"]
    empty_value5=empty_list.at[i,"level5"]
    for j in range(len(npt)):
        if empty_index==j:
            npt.at[j,"level1"]=empty_value1
            npt.at[j,"level2"]=empty_value2
            npt.at[j,"level3"]=empty_value3
            npt.at[j,"level4"]=empty_value4
            npt.at[j,"level5"]=empty_value5
        else:
            pass

#비교할 NPT 데이터만 남기기 (Assembly Pull Type)
npt=npt[npt['Supply Type']=='Assembly Pull']
npt.reset_index(inplace=True,drop=True)

npt.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/nptresult.xlsx')

####################### compare npt and gerp #######################
#Match with GERP parent
#Change Parent Part As F3 (NPT and GERP match) remove TAW
f3_parent=gerp.at[0,"Parent Item"]
for i in range(len(npt)):
    parent_part=npt.at[i,'Parent Part'][:3]
    if parent_part=="TAW":
        npt.at[i,'Parent Part']=f3_parent
    else:
        pass

#match_list 시작
match_list=pd.DataFrame()
unique_gerp=gerp
model_name=gerp.at[0,"Parent Item"][:1]

#Screw,Tapping / Screw,Taptite / Duct / Packing / Resin,EPS / Damper 우선 배치 혼돈을 줄수 있음
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_parent=npt.at[i,"Parent Part"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_qty=npt.at[i,"Unit Qty"]
    match_number=npt.at[i,"Seq."]
    if npt_des=="Screw,Tapping": #Screw,tapping 
        if model_name=='T' or 'F': #Screw,tapping -> TL,FL
            for j in range(len(unique_gerp)): #duplicate 제외
                gerp_part=unique_gerp.at[j,"Child Item"]
                gerp_des=unique_gerp.at[j,"Description"]
                gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
                gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
                if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty: #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
                    if gerp_price==npt_price:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_true"]=gerp_data
                    else:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_price"]=gerp_data
                else:
                    continue
            ### HOW TO DROP 
            used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
            unique_gerp=unique_gerp.drop(used_index,axis=0)
            unique_gerp.reset_index(drop=True, inplace=True)

        elif model_name=='R': #Screw,tapping -> TL,FL
            for j in range(len(unique_gerp)): #duplicate 제외
                gerp_parent=unique_gerp.at[j,"Parent Item"]
                gerp_part=unique_gerp.at[j,"Child Item"]
                gerp_des=unique_gerp.at[j,"Description"]
                gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
                gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
                if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty and gerp_parent==npt_parent: #완전 일치문 (parent는 일치해야 함 드라이어)
                    if gerp_price==npt_price:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_true"]=gerp_data
                    else:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_price"]=gerp_data
                else:
                    continue
            ### HOW TO DROP 
            used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
            unique_gerp=unique_gerp.drop(used_index,axis=0)
            unique_gerp.reset_index(drop=True, inplace=True)
        else:
            print("Unidentified Model")
    


    #Screw,Taptite -> DR
    elif npt_des=='Screw,Taptite':
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty: #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Duct Assembly -> DR
    elif npt_des=='Duct Assembly':
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_sub_part=str(unique_gerp.at[j,"Child Item"])[:-1] # 끝자리 한자리까지 일치
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            if npt_part.__contains__(gerp_sub_part): # 끝자리 한자리까지 일치
                if gerp_part==npt_part: #완전일치
                    if gerp_price==npt_price:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_true"]=gerp_data
                    else:
                        gerp_data=unique_gerp.at[j,"Seq"]
                        match_list.at[match_number,"gerp_price"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_sub"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Packing -> DR 
    elif npt_des=='Packing':
        for j in range(len(unique_gerp)):
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part: #부모랑 des일치하면 우선 배정, 파트넘버 전부 같음
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)

    #Resin,EPS -> DR 
    elif npt_des=='Resin,EPS':
        for j in range(len(unique_gerp)):
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and npt_parent==gerp_parent: #부모랑 des일치하면 우선 배정, 파트넘버 전부 같음
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)
    else:
        pass

#Screw,Customized -> 부모, description, Qty 만보고 먼저 매치
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_qty=npt.at[i,"Unit Qty"]
    match_number=npt.at[i,"Seq."]
    if npt_des=="Screw,Customized":
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            gerp_qty=unique_gerp.at[j,"Qty Per Assembly"]
            if gerp_des==npt_des and gerp_part==npt_part and gerp_qty==npt_qty: #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
                if gerp_price==npt_price:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_true"]=gerp_data
                else:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_price"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)
    else:
        pass


#Hanger,Upper -> 부모, description 만보고 먼저 매치
sub_count=0
count=0
gerp_drop=pd.DataFrame()
for i in range(len(npt)):
    npt_des=npt.at[i,"Desc."]
    npt_part=npt.at[i,"Part No"]
    npt_sub=str(npt.at[i,"Part No"])[:-2] #part넘버가 모두 숫자인 경우
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_sub=str(unique_gerp.at[j,"Child Item"])[:-2] #part넘버가 모두 숫자인 경우
        #완전 일치문 (parent는 일치하지 않는 경우가 있어 제외)
        if gerp_des=='Hanger,Upper' and gerp_des==npt_des and gerp_part==npt_part: 
            if gerp_price==npt_price:
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_true"]=unique_gerp.at[j,"Seq"] #gerp는 j
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_numbre=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_price"]=unique_gerp.at[j,"Seq"] #gerp는 j
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        #sub로 매치
        elif gerp_des=='Hanger,Upper' and gerp_des==npt_des and gerp_sub==npt_sub: 
            if sub_count==0: #한번도 sub 매치 안된 경우
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_sub"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=sub_count+1
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_exc"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=0
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        #door glass -> parent 다른게 두게
        elif gerp_des=="Door,Glass" and npt_des==gerp_des and gerp_sub==npt_sub:
            if gerp_part==npt_part: #parent 다르지만 part 일치
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_price"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=sub_count+1
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1
            else:
                match_number=npt.at[i,"Seq."]
                match_list.at[match_number,"gerp_exc"]=unique_gerp.at[j,"Seq"] #gerp는 j
                sub_count=0
                #드롭할 것 담기
                gerp_drop.at[count,"gerp"]=unique_gerp.at[j,"Seq"]
                count=count+1

        else:
            continue

### HOW TO DROP 
unique_gerp.index=unique_gerp["Seq"] #index reset
if len(gerp_drop)==0:
    unique_gerp.reset_index(drop=True, inplace=True)
else:
    drop_index=gerp_drop.gerp.tolist()
    unique_gerp=unique_gerp.drop(drop_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)



# Holder -> 파트넘버, des 보고 바로 매치
for i in range(len(npt)):
    npt_parent=npt.at[i,"Parent Part"]
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    npt_level=int(npt.at[i,"Lvl"][-1:])
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if npt_des=="Holder" and npt_level>2 and npt_des==gerp_des and npt_part==gerp_part:
            if gerp_price==npt_price:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_true"]=gerp_data
            else:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_price"]=gerp_data
        else:
            continue
    else:
        pass
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)

#일치 조건문 - 완전일치
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part==npt_part and gerp_des==npt_des and gerp_parent==npt_parent:#완전 일치문
            if gerp_price==npt_price:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_true"]=gerp_data
            else:
                gerp_data=unique_gerp.at[j,"Seq"]
                match_list.at[match_number,"gerp_price"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)


#일치 조건문 - substitute   (part no 2자리 빼고 똑같은 것 ==> 둘다 parent와 description)
for i in range(len(npt)):
    npt_part=str(npt.at[i,"Part No"])
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=str(unique_gerp.at[j,"Child Item"])
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part[:-2]==npt_part[:-2] and gerp_des==npt_des and gerp_parent==npt_parent:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_sub"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True)    


#일치 조건문 - substitute   (part no 완전 다른 것 ==> 둘다 parent와 description)
for i in range(len(npt)):
    #Label,Barcode의 경우 매칭되지 않도록 주의 
    npt_des=npt.at[i,"Desc."]
    if npt_des=='Label,Barcode': #남는 데이터로 다시 돌리는 조건문에서 매칭 될 것
        pass
    else:
        npt_part=npt.at[i,"Part No"]
        npt_des=npt.at[i,"Desc."]
        npt_parent=npt.at[i,"Parent Part"]
        npt_price=npt.at[i,"Material Cost (LOC)"]
        match_number=npt.at[i,"Seq."]
        for j in range(len(unique_gerp)): #duplicate 제외
            gerp_part=unique_gerp.at[j,"Child Item"]
            gerp_des=unique_gerp.at[j,"Description"]
            gerp_parent=unique_gerp.at[j,"Parent Item"]
            gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
            if gerp_des==npt_des and gerp_parent==npt_parent:
                if gerp_part!=npt_part:
                    gerp_data=unique_gerp.at[j,"Seq"]
                    match_list.at[match_number,"gerp_exc"]=gerp_data
            else:
                continue
        ### HOW TO DROP 
        used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
        unique_gerp=unique_gerp.drop(used_index,axis=0)
        unique_gerp.reset_index(drop=True, inplace=True)   


#일치 조건문 - parent 불일치
for i in range(len(npt)):
    npt_part=npt.at[i,"Part No"]
    npt_des=npt.at[i,"Desc."]
    npt_parent=npt.at[i,"Parent Part"]
    npt_price=npt.at[i,"Material Cost (LOC)"]
    match_number=npt.at[i,"Seq."]
    for j in range(len(unique_gerp)): #duplicate 제외
        gerp_part=unique_gerp.at[j,"Child Item"]
        gerp_des=unique_gerp.at[j,"Description"]
        gerp_parent=unique_gerp.at[j,"Parent Item"]
        gerp_price=unique_gerp.at[j,"QPA*Material Cost"]
        if gerp_part==npt_part and gerp_des==npt_des:
            gerp_data=unique_gerp.at[j,"Seq"]
            match_list.at[match_number,"gerp_parent"]=gerp_data
        else:
            continue
    ### HOW TO DROP 
    used_index=unique_gerp[unique_gerp["Seq"]==gerp_data].index
    unique_gerp=unique_gerp.drop(used_index,axis=0)
    unique_gerp.reset_index(drop=True, inplace=True) 


### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
#price, parent -> parent는 드롭 -> unique_add
match_list=match_list.reset_index()
unique_parent=pd.DataFrame()
count=0
for i in range(len(match_list)):
    price=str(match_list.at[i,"gerp_price"])
    parent=str(match_list.at[i,"gerp_parent"])
    if price!='nan' and parent!='nan':
        unique_parent.at[count,"Seq"]=match_list.at[i,"gerp_parent"]
        count=count+1
    else:
        pass

#unique_parent 존재 유무
unique_parent_count=len(unique_parent)
if unique_parent_count==0:
    remain_gerp=unique_gerp.dropna(subset=['Seq']) #seq빈 애들 drop
    remain_gerp.reset_index(drop=True, inplace=True)
else:
    A=unique_parent["Seq"].tolist()
    unique_add=gerp
    unique_add.index=unique_add.Seq
    unique_add=unique_add.iloc[A]
    unique_add.index=range(len(unique_gerp), len(unique_gerp)+len(unique_add))
    #unique_add 를 unique_gerp에 더하기 & 내림 차순
    remain_gerp=pd.concat([unique_gerp, unique_add], axis=0).sort_values(by=['Seq'],ascending=True)
    remain_gerp.reset_index(drop=True, inplace=True)
    #남은 것 다시 매칭하기 위해 행렬 정렬 -> 아래 조건문에서 자기 자리 찾아감
    match_list.index=match_list['index']


### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
match_list.index=match_list['index']

#남은 것 다시 매칭--> Coil,Steel => substitue part인 경우
match_list.index=match_list['gerp_price'] # gerp에서 부모가 같은 경우, Handle를 포함
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=="Coil,Steel(STS)":
        for j in range(len(gerp)):
            gerp_parent=gerp.at[j,"Parent Item"]
            gerp_des=gerp.at[j,"Description"]
            gerp_seq=gerp.at[j,"Seq"]
            if gerp_des=="Sheet,Steel(STS)" and gerp_parent==remain_parent:
                match_list.at[gerp_seq,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
        else:
            pass

match_list.reset_index(drop=True, inplace=True)
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)

#남은 것 다시 매칭--> pulsator cover/steel sheet/duct
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_part=remain_gerp.at[i,"Child Item"][:-2]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    remain_sub_parent=str(remain_gerp.at[i,"Parent Item"])[:-2]
    for j in range(len(npt)):
        #hanger pivot -> 2 gerp_parent
        npt_des=npt.at[j,"Desc."]
        npt_part=str(npt.at[j,"Part No"])[:-2]
        npt_parent=str(npt.at[j,"Parent Part"])
        #Match to npt-pulsator cover
        if remain_des=='Cover,Pulsator':
            if remain_des==npt_des and remain_part==npt_part:
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass
        
        #Match to npt-Coil,Steel(GI -> DR)
        elif remain_des=='Coil,Steel(GI)':
            if remain_parent==npt_parent and npt_des=='Sheet,Steel(GA)':
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            elif remain_part==npt_part and npt_des==remain_des:
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass

        #Match to npt-sheet,Steel(GI -> DR)
        elif remain_des=='Sheet,Steel(GI)':
            if remain_parent==npt_parent and npt_des=='Coil,Steel(GI)':
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass
        else:
            pass
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


#남은 것 다시 매칭--> resin,asa
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"][:-2]
    remain_part=remain_gerp.at[i,"Child Item"][:-2]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    for j in range(len(npt)):
        #hanger pivot -> 2 gerp_parent
        npt_des=npt.at[j,"Desc."][:-2]
        npt_part=str(npt.at[j,"Part No"])[:-2] #substitute로 같은 애도 인정
        npt_parent=npt.at[j,"Parent Part"]
        if npt_des=='Resin,A' and npt_parent==remain_parent:
            match_number=npt.at[j,"Seq."]
            match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
            remain_match.at[count,"index"]=i
            count=count+1
        else:
            pass
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


#남은 것 다시 매칭--> coil,sheet(sts)
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    for j in range(len(npt)):
        #hanger pivot -> 2 gerp_parent
        npt_des=npt.at[j,"Desc."]
        if npt_des=='Sheet,Steel(STS)' and remain_des=='Coil,Steel(STS)':
            match_number=npt.at[j,"Seq."]
            match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
            remain_match.at[count,"index"]=i
            count=count+1
        else:
            pass
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


#남은 것 다시 매칭--> stator, rotor combined
remain_match=pd.DataFrame()
match_list.index=match_list['index']
count=0
for i in range(len(remain_gerp)): #i -> gerp
    #stator, rotor combined
    remain_des=remain_gerp.at[i,"Description"]
    if remain_des.__contains__('Assembly,Combined'):
        for j in range(len(npt)): #j -> npt
            npt_des=npt.at[j,"Desc."]
            if remain_des.__contains__(npt_des):
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass
    
    #rotor assembly  -> FL
    elif remain_des=='Rotor Assembly':
        for j in range(len(npt)):
            npt_des=npt.at[j,"Desc."]
            if npt_des.__contains__('Rotor'):
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass

    #stator assembly  -> FL
    elif remain_des=='Stator Assembly':
        for j in range(len(npt)):
            npt_des=npt.at[j,"Desc."]
            if npt_des.__contains__('Stator'):
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass

    #나머지는 통과
    else:
        pass

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
#남은 것 다시 매칭--> hanger pivot part => gerp_parent 값이 두개인 경우
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_part=remain_gerp.at[i,"Child Item"][:-2]
    for j in range(len(npt)):
        #hanger pivot -> 2 gerp_parent
        npt_des=npt.at[j,"Desc."]
        npt_part=str(npt.at[j,"Part No"])[:-2]
        if remain_des.__contains__(',Pivot'):
            if remain_des==npt_des and remain_part==npt_part:
                match_number=npt.at[j,"Seq."]
                match_list.at[match_number,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
            else:
                pass
        else:
            pass
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
#남은 것 다시 매칭--> handle assembly => gerp_parent 값이 두개인 경우
match_list.index=match_list['gerp_parent'] # gerp에서 부모가 같은 경우, Handle를 포함
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=='Handle':
        for j in range(len(gerp)):
            gerp_parent=gerp.at[j,"Parent Item"]
            gerp_des=gerp.at[j,"Description"]
            gerp_seq=gerp.at[j,"Seq"]
            if gerp_des=='Handle Assembly' and gerp_parent==remain_parent:
                match_list.at[gerp_seq,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
    elif remain_des=='Handle Assembly':
        for j in range(len(gerp)):
            gerp_parent=gerp.at[j,"Parent Item"]
            gerp_des=gerp.at[j,"Description"]
            gerp_seq=gerp.at[j,"Seq"]
            if gerp_des=='Handle' and gerp_parent==remain_parent:
                match_list.at[gerp_seq,"gerp_re"]=remain_gerp.at[i,"Seq"]
                remain_match.at[count,"index"]=i
                count=count+1
        else:
            pass
match_list.reset_index(drop=True, inplace=True)
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


### 매칭 안된 데이터 - 남는 데이터 가지고 다시 비교
#남은 것 다시 매칭--> door glass => gerp_parent 값이 두개인 경우
match_list.index=match_list['gerp_parent'] # gerp_parent를 인덱스로 두고 포인팅 값이 표기
remain_match=pd.DataFrame()
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=remain_gerp.at[i,"Description"]
    remain_parent=remain_gerp.at[i,"Parent Item"]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=='Door,Glass':
        for k in range(len(gerp)):
            gerp_parent=gerp.at[k,"Parent Item"]
            gerp_des=gerp.at[k,"Description"]
            gerp_seq=gerp.at[k,"Seq"]
            if gerp_parent==remain_parent and gerp_des=='Door,Glass':
                match_list.at[k,"gerp_re"]=remain_seq

                # 빈 데이터 프레임에 값 저장 -> to drop on the remain_gerp list
                remain_match.at[count,"index"]=i #remain_gerp list index
                count=count+1
        else:
            pass
#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    match_list.reset_index(inplace=True, drop=True)
    remain_gerp.reset_index(inplace=True, drop=True)
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)
    match_list.reset_index(inplace=True, drop=True)

### Remain gerp : Label,Barcode -> FL
remain_match=pd.DataFrame()
match_list.index=match_list["index"]
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=str(remain_gerp.at[i,"Description"])
    remain_part=remain_gerp.at[i,"Child Item"][:-2]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=="Label,Barcode":
        for j in range(len(gerp)):
            npt_part=str(npt.at[j,"Part No"])[:-2]
            npt_seq=npt.at[j,"Seq."]
            npt_des=npt.at[j,"Desc."]
            if npt_part==remain_part and npt_des==remain_des:
                match_list.at[npt_seq,"gerp_sub"]=remain_seq
                match_list.at[npt_seq,"index"]=npt_seq
                # 빈 데이터 프레임에 값 저장 -> to drop on the remain_gerp list
                remain_match.at[count,"index"]=i #remain_gerp list index
                count=count+1

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    remain_gerp.reset_index(inplace=True, drop=True)
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)

#remain gerp drum
remain_match=pd.DataFrame()
match_list.index=match_list["index"]
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=str(remain_gerp.at[i,"Description"])
    remain_parent=remain_gerp.at[i,"Parent Item"][:-2]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=="Coil,Steel(STS)":
        for j in range(len(gerp)):
            npt_parent=str(npt.at[j,"Parent Part"])[:-2]
            npt_seq=npt.at[j,"Seq."]
            npt_des=npt.at[j,"Desc."]
            if npt_parent==remain_parent and npt_des==remain_des:
                match_list.at[npt_seq,"gerp_sub"]=remain_seq
                match_list.at[npt_seq,"index"]=npt_seq
                # 빈 데이터 프레임에 값 저장 -> to drop on the remain_gerp list
                remain_match.at[count,"index"]=i #remain_gerp list index
                count=count+1

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    remain_gerp.reset_index(inplace=True, drop=True)
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)


#remain gerp rear cover 
remain_match=pd.DataFrame()
match_list.index=match_list["index"]
count=0
for i in range(len(remain_gerp)): #i -> gerp
    remain_des=str(remain_gerp.at[i,"Description"])
    remain_parent=remain_gerp.at[i,"Parent Item"][:-2]
    remain_seq=remain_gerp.at[i,"Seq"]
    if remain_des=="Sheet,Steel(GI)":
        for j in range(len(gerp)):
            npt_child=str(npt.at[j,"Part No"])[:-2]
            npt_seq=npt.at[j,"Seq."]
            npt_des=npt.at[j,"Desc."]
            if remain_parent==npt_child and npt_des=='Cover,Rear':
                match_list.at[npt_seq,"gerp_exc"]=remain_seq
                match_list.at[npt_seq,"index"]=npt_seq
                # 빈 데이터 프레임에 값 저장 -> to drop on the remain_gerp list
                remain_match.at[count,"index"]=i #remain_gerp list index
                count=count+1

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    remain_gerp.reset_index(inplace=True, drop=True)
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)

#reset index of match_list to save final file
match_list.reset_index(drop=True, inplace=True)

#매칭 안되고 missing 된것
remain_gerp.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/remaingerp.xlsx')

#gerp_re할때 순서 바꾼 것 다시 원상 복귀, 매칭 리스트 파일 위해 인덱스 재정렬
match_list.reset_index(inplace=True, drop=True)

#matchlist에서 count만들기
match_list['count']=match_list.count(axis = 1)-1
match_list=match_list.rename(columns = {'index':'Seq.'})

#save file
match_list.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/matchlist.xlsx')

#matchlist에서 substitute part가 있는 애들은 칼럼 두개 만들기
sub_match_list=pd.DataFrame()
change_count=0

#칼럼의 판별을 위해
match_column=match_list.columns

if match_column.__contains__('gerp_true'): #trueO
    if match_column.__contains__('gerp_re'): #trueO reO
        if match_column.__contains__('gerp_exc'): # trueO reO excO
            true_exist=True
            re_exist=True
            exc_exist=True
        else: # trueO reO excX
            true_exist=True
            re_exist=True
            exc_exist=False

    else: #trueO reX
        if match_column.__contains__('gerp_exc'): # trueO reX excO
            true_exist=True
            re_exist=False
            exc_exist=True
        else: # trueO reX excX
            true_exist=True
            re_exist=False
            exc_exist=False          
else: #trueX
    if match_column.__contains__('gerp_re'): # trueX reO
        if match_column.__contains__('gerp_exc'): # trueX reO excO
            true_exist=False
            re_exist=True
            exc_exist=True
        else: # trueX reO excX
            true_exist=False
            re_exist=True
            exc_exist=False
    else: #trueX reX
        if match_column.__contains__('gerp_exc'): #trueX reX excO
            true_exist=False 
            re_exist=False
            exc_exist=True
        else: #trueX reX excX
            true_exist=False
            re_exist=False
            exc_exist=False


#submatchlist
for i in range(len(match_list)):
    ##################################### gerp_re column exist
    if re_exist==True: 
        match_count=match_list.at[i,"count"]
        sub_empty=str(match_list.at[i,"gerp_sub"])
        price_empty=str(match_list.at[i,"gerp_price"])
        parent_empty=str(match_list.at[i,"gerp_parent"])
        #gerp_re exist or not
        if re_exist==True:
            re_empty=str(match_list.at[i,"gerp_re"])
        else:
            re_empty='nan'

        #gerp_exc column exist
        if exc_exist==True:
            exc_empty=str(match_list.at[i,"gerp_exc"])
        else:
            exc_empty='nan'

        #value
        sub_part=match_list.at[i,"gerp_sub"]
        price_part=match_list.at[i,"gerp_price"]
        if exc_exist==True:
            exc_part=match_list.at[i,"gerp_exc"]
        else:
            exc_part='nan'

        ############################ sub value exist ==> 한칸 뒤로 뺄 필요 (sub 칼럼은 항상 존재)
        if sub_empty!='nan': 
            ######sub, price exist
            if price_empty!='nan': 
                #######sub, price, exc exist
                if exc_empty!='nan':
                    ##### exc exist, price exist, sub exist
                    ### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part # 채움
                    sub_match_list.at[change_count,'gerp_sub']=np.nan # sub 제거 
                    sub_match_list.at[change_count,'gerp_exc']=np.nan # exc 제거
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exist, price, sub exist -> 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

                    ### only sub column
                    change_count=change_count+1 # 한칸뒤로
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price 제거
                    sub_match_list.at[change_count,'gerp_sub']=sub_part # 채움
                    sub_match_list.at[change_count,'gerp_exc']=np.nan # exc 제거 
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price, sub exist -> 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

                    ### only exc column 
                    change_count=change_count+1 #한칸뒤로
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=np.nan # sub 제거
                    sub_match_list.at[change_count,'gerp_exc']=exc_part #채움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price, sub exist -> 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)
                
                else:
                #######sub, price exist
                #### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part #채움
                    sub_match_list.at[change_count,'gerp_sub']=np.nan #sub제거
                    if exc_empty!='nan':
                        sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상 (조건문)
                    else:
                        pass
                    sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist, price exist -> parent 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)
                    
                    #### only sub column 
                    change_count=change_count+1
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=sub_part #채움
                    if exc_empty!='nan':
                        sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상 (조건문)
                    else:
                        pass
                    sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist, price exist -> parent 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

            else: 
                ##### sub exist, price existX, re existX
                ### only sub column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=np.nan # 조건문에서 아닌 것 판명
                sub_match_list.at[change_count,'gerp_sub']=sub_part #채움
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상
                sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist -> parent 필요 없음
                sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

        ##################### exc value exist ==> 한칸 뒤로 뺄 필요
        elif exc_empty!='nan': 
            #exc part, price part
            price_part=match_list.at[i,"gerp_price"]
            exc_part=match_list.at[i,"gerp_exc"]
            ########### exc, price value exist
            if price_empty!='nan': 
                ########### exc, price value exist / sub value exist X 
                if sub_empty=='nan':
                    ### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part #채움
                    sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_exc']=np.nan #비움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price exist -> 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)
                    
                    ### only exc column 
                    change_count=change_count+1
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_exc']=exc_part # 채움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price exist -> 필요 없음
                    sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)
                else:
                    print("Logic Error")

            #re,parent,exc
            elif re_empty!='nan':
                ########### exc, price, re exist
                ### only parent column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"] #빔 예상 (추측)
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (추측)
                sub_match_list.at[change_count,'gerp_exc']=np.nan #비움
                sub_match_list.at[change_count,'gerp_parent']=match_list.at[i,"gerp_parent"] #채움
                sub_match_list.at[change_count,'gerp_re']=np.nan #비움
                
                ### only re column 
                change_count=change_count+1
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=np.nan #비움
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_exc']=np.nan #비움
                sub_match_list.at[change_count,'gerp_parent']=np.nan #비움
                sub_match_list.at[change_count,'gerp_re']=match_list.at[i,"gerp_re"] #채움

                ### only exc column 
                change_count=change_count+1
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=np.nan #비움
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] # 채움
                sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price exist -> 필요 없음
                sub_match_list.at[change_count,'gerp_re']=np.nan #비움
            else: 
                ########### exc exist
                #only exc column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"] #빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상
                sub_match_list.at[change_count,'gerp_exc']=exc_part #채움
                sub_match_list.at[change_count,'gerp_parent']=np.nan # exc exist-> 필요 없음
                sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

        ##################### re value exist ==> 한칸 뒤로 뺄 필요
        elif re_empty!='nan': 
            #exc part, price part
            price_part=str(match_list.at[i,"gerp_price"])
            parent_part=str(match_list.at[i,"gerp_parent"])
            re_part=str(match_list.at[i,"gerp_re"])

            #re, price 존재
            if price_part!='nan':
                ### only price column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=price_part #채움
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]#빔 예상 (조건문)
                if exc_empty!='nan':
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]#빔 예상 (조건문)
                else:
                    pass
                sub_match_list.at[change_count,'gerp_parent']=match_list.at[i,"gerp_parent"]#빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_re']=np.nan #비움
                
                ### only re column 
                change_count=change_count+1
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=np.nan#비움
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]#빔 예상 (조건문)
                if exc_empty!='nan':
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]#빔 예상 (조건문)
                else:
                    pass
                sub_match_list.at[change_count,'gerp_parent']=match_list.at[i,"gerp_parent"]#빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_re']=re_part #채움

            #re,parent
            elif parent_part!='nan':
                ### only parent column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]#빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]#빔 예상 (조건문)
                if exc_empty!='nan':
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]#빔 예상 (조건문)
                else:
                    pass
                sub_match_list.at[change_count,'gerp_parent']=parent_part #채움:
                sub_match_list.at[change_count,'gerp_re']=np.nan #비움
                
                ### only re column 
                change_count=change_count+1
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                # true column
                if true_exist==True: # true exist
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true exist
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)")
                #other column
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]#빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]#빔 예상 (조건문)
                if exc_empty!='nan':
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]#빔 예상 (조건문)
                else:
                    pass
                sub_match_list.at[change_count,'gerp_parent']=np.nan #비움
                sub_match_list.at[change_count,'gerp_re']=re_part #채움
            else:
                pass

        else: 
            ###################  sub empty and exc empty ==> 한칸 뒤로 뺄 필요 없음
            if parent_empty!='nan' and price_empty!='nan':
            #### price, parent 둘다 있는 경우 price 선택
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                #other column
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] # 빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] # 빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_parent']=np.nan #parent 버리기
                sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

            else: 
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]
                if exc_empty!='nan':
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]
                else:
                    pass
                sub_match_list.at[change_count,'gerp_parent']=match_list.at[i,"gerp_parent"]
                sub_match_list.at[change_count,'gerp_re']=match_list.at[i,'gerp_re'] #빔 예상(조건문)

        #사이클 돌리기 위해 +1 증가
        change_count=change_count+1
    
    
    ##################################### gerp_re column exist X
    else:
        ##################################### gerp_re exist X
        match_count=match_list.at[i,"count"]
        sub_empty=str(match_list.at[i,"gerp_sub"])
        price_empty=str(match_list.at[i,"gerp_price"])
        exc_empty=str(match_list.at[i,"gerp_exc"])
        parent_empty=str(match_list.at[i,"gerp_parent"])

        #value
        sub_part=match_list.at[i,"gerp_sub"]
        price_part=match_list.at[i,"gerp_price"]
        exc_part=match_list.at[i,"gerp_exc"]

        if sub_empty!='nan': 
            ############################ sub exist ==> 한칸 뒤로 뺄 필요
            ######sub, price exist
            if price_empty!='nan': 
                #######sub, price, exc exist
                if exc_empty!='nan':
                    ##### exc exist, price exist, sub exist
                    ### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part # 채움
                    sub_match_list.at[change_count,'gerp_sub']=np.nan # sub 제거 
                    sub_match_list.at[change_count,'gerp_exc']=np.nan # exc 제거
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exist, price, sub exist -> 필요 없음

                    
                    ### only sub column
                    change_count=change_count+1 # 한칸뒤로
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price 제거
                    sub_match_list.at[change_count,'gerp_sub']=sub_part # 채움
                    sub_match_list.at[change_count,'gerp_exc']=np.nan # exc 제거 
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price, sub exist -> 필요 없음


                    ### only exc column 
                    change_count=change_count+1 #한칸뒤로
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=np.nan # sub 제거
                    sub_match_list.at[change_count,'gerp_exc']=exc_part #채움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price, sub exist -> 필요 없음

                else:
                #######sub, price exist
                #### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part #채움
                    sub_match_list.at[change_count,'gerp_sub']=np.nan #sub제거
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist, price exist -> parent 필요 없음

                    
                    #### only sub column 
                    change_count=change_count+1
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=sub_part #채움
                    sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist, price exist -> parent 필요 없음

            else: 
                ##### sub exist, price existX, re existX
                ### only sub column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=np.nan # 조건문에서 아닌 것 판명
                sub_match_list.at[change_count,'gerp_sub']=sub_part #채움
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] #빔 예상
                sub_match_list.at[change_count,'gerp_parent']=np.nan #sub exist -> parent 필요 없음


        elif exc_empty!='nan': 
            ##################### exc exist ==> 한칸 뒤로 뺄 필요
            #exc part, price part
            price_part=match_list.at[i,"gerp_price"]
            exc_part=match_list.at[i,"gerp_exc"]
            if price_empty!='nan': 
                ########### exc exist, price exist
                if sub_empty=='nan':
                    ########### exc, price exist and sub exist X 
                    ### only price column
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=price_part #채움
                    sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_exc']=np.nan #비움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price exist -> 필요 없음


                    ### only exc column 
                    change_count=change_count+1
                    sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                    # true column
                    if true_exist==True: # true exist
                        sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                    elif true_exist==False: # true exist
                        pass
                    else:
                        print("Except True column there is another column missing (price, sub, parent)")
                    #other column
                    sub_match_list.at[change_count,"gerp_price"]=np.nan # price제거
                    sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상 (조건문)
                    sub_match_list.at[change_count,'gerp_exc']=exc_part # 채움
                    sub_match_list.at[change_count,'gerp_parent']=np.nan # exc, price exist -> 필요 없음
                
                else:
                    print("no logic")


            else: 
                ########### exc exist
                #only exc column
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"] #빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] #빔 예상
                sub_match_list.at[change_count,'gerp_exc']=exc_part #채움
                sub_match_list.at[change_count,'gerp_parent']=np.nan # exc exist-> 필요 없음

        else: 
            ###################  sub empty and exc empty ==> 한칸 뒤로 뺄 필요 없음
            if parent_empty!='nan' and price_empty!='nan':
            #### price, parent 둘다 있는 경우 price 선택
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                #other column
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"] # 빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"] # 빔 예상 (조건문)
                sub_match_list.at[change_count,'gerp_parent']=np.nan #parent 버리기

            else: 
                sub_match_list.at[change_count,"Seq."]=match_list.at[i,"Seq."]
                if true_exist==True: # true 존재O
                    sub_match_list.at[change_count,"gerp_true"]=match_list.at[i,"gerp_true"]
                elif true_exist==False: # true 존재X
                    pass
                else:
                    print("Except True column there is another column missing (price, sub, parent)") 
                sub_match_list.at[change_count,"gerp_price"]=match_list.at[i,"gerp_price"]
                sub_match_list.at[change_count,'gerp_sub']=match_list.at[i,"gerp_sub"]
                sub_match_list.at[change_count,'gerp_exc']=match_list.at[i,"gerp_exc"]
                sub_match_list.at[change_count,'gerp_parent']=match_list.at[i,"gerp_parent"]

        #사이클 돌리기 위해 +1 증가
        change_count=change_count+1    
    
#save file
sub_match_list.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/submatchlist.xlsx')

#아무것도 해당안되는 애들 뽑아서 sub_match_list에 넣기
match_join = pd.merge(npt['Seq.'], sub_match_list, on ='Seq.', how ='left')

#nan -> count
match_join['count']=match_join.count(axis=1)-1
#nan -> 0(더하기 위한)
match_sum=match_join.fillna(0)
#gerpSeq. create
if true_exist==True: # true 존재O
    if re_exist==True: #re 존재O
        if exc_exist==True: #exc 존재O
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_true'])+pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_exc'])+pd.to_numeric(match_sum['gerp_parent'])+pd.to_numeric(match_sum['gerp_re'])
        else:
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_true'])+pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_parent'])+pd.to_numeric(match_sum['gerp_re'])
    else:
        if exc_exist==True: #exc 존재O
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_true'])+pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_exc'])+pd.to_numeric(match_sum['gerp_parent'])
        else:
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_true'])+pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_parent'])
  

else: 
    if re_exist==True: #re 존재O
        if exc_exist==True: #exc 존재O
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_exc'])+pd.to_numeric(match_sum['gerp_parent'])+pd.to_numeric(match_sum['gerp_re'])
        else:
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_parent'])+pd.to_numeric(match_sum['gerp_re'])
          
    else: 
        if exc_exist==True: #exc 존재O
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_exc'])+pd.to_numeric(match_sum['gerp_parent'])
        else:
            match_join['gerpSeq.']=pd.to_numeric(match_sum['gerp_price'])+pd.to_numeric(match_sum['gerp_sub'])+pd.to_numeric(match_sum['gerp_parent'])
           

#count=2전부1로 만들기
for i in range(len(match_join)):
    match_count=match_join.at[i,"count"]
    match_price=match_join.at[i,"gerp_price"]
    if match_count==2:
        if match_price!=np.nan: #값이 있음
            match_join.at[i,'gerpSeq.']=match_join.at[i,'gerpSeq.']-match_join.at[i,'gerp_parent']
            match_join.at[i,'gerp_parent']=np.nan #''로 하면 카운트 됨
            match_join.at[i,'count']=1 #''로 하면 카운트 됨
        else:
            print("gerp_price, gerp_parent only - Error") #gerp_price, gerp_parent 이외에 다른 게 count2를 만듬

#save file
match_join.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/matchjoin.xlsx')

#final table
final_table=pd.DataFrame()
match_column=pd.DataFrame()
for i in range(len(match_join)):
    #substitute part npt delete
    npt_seq=match_join.at[i,"Seq."]
    gerp_seq=match_join.at[i,"gerpSeq."]

    #NPT
    #sequence 일치 칼럼을 뽑아내고
    npt_column=npt[npt["Seq."]==npt_seq]
    npt_column.reset_index(inplace=True, drop=True) #index 넘버 상관 X하고 0으로 
    #npt list
    final_table.at[i,"Seq."]=npt_seq
    final_table.at[i,"level"]=npt_column.at[0,"level"]
    final_table.at[i,"Parent Part"]=npt_column.at[0,"Parent Part"]
    final_table.at[i,"Part No"]=npt_column.at[0,"Part No"]
    final_table.at[i,"Desc."]=npt_column.at[0,"Desc."]
    final_table.at[i,"UOM"]=npt_column.at[0,"UOM"]
    final_table.at[i,"Unit Qty"]=npt_column.at[0,"Unit Qty"]
    final_table.at[i,"Accum Qty"]=npt_column.at[0,"Accum Qty"]
    final_table.at[i,"Supplier Name"]=npt_column.at[0,"Supplier Name"]
    final_table.at[i,"Curr."]=npt_column.at[0,"Curr."]
    final_table.at[i,"Unit Price"]=npt_column.at[0,"Unit Price"]
    final_table.at[i,"Exchange Rate(LOC)"]=npt_column.at[0,"Exchange Rate(LOC)"]
    final_table.at[i,"Unit Price (USD)"]=npt_column.at[0,"Unit Price (USD)"]
    final_table.at[i,"Material Cost (LOC)"]=npt_column.at[0,"Material Cost (LOC)"]
    #match list
    final_table.at[i,"npt end"]=''
    final_table.at[i,"match"]=''
    final_table.at[i,"price match"]=''
    final_table.at[i,"gerp start"]=''

    #GERP Matching
    match_count=match_join.at[i,"count"]
    match_sub=match_join.at[i,"gerp_sub"]
    #GERP ==> Not Match
    if match_count==0: #empty dataframe -> count0
        final_table.at[i,"Seq"]=''
        final_table.at[i,"Level"]=''
        final_table.at[i,"Parent Item"]=''
        final_table.at[i,"Child Item"]=''
        final_table.at[i,"Description"]=''
        final_table.at[i,"Item Type"]=''
        final_table.at[i,"Qty Per Assembly"]=''
        final_table.at[i,"Material Cost"]=''
        final_table.at[i,"QPA*Material Cost"]=''
        #match -> not match
        final_table.at[i,'match']="False"
        #나중에 채울 price match
        final_table.at[i,"price match"]=''
    #GERP ==> Match
    else:
        #fill dataframe -> count1
        gerp_column=gerp[gerp["Seq"]==gerp_seq]
        gerp_column.reset_index(inplace=True, drop=True) #index 넘버 상관 X하고 0으로 
        final_table.at[i,"Seq"]=gerp_seq
        final_table.at[i,"Level"]=int(str(gerp_column.at[0,"Level"])[-1:])
        final_table.at[i,"Parent Item"]=gerp_column.at[0,"Parent Item"]
        final_table.at[i,"Child Item"]=gerp_column.at[0,"Child Item"]
        final_table.at[i,"Description"]=gerp_column.at[0,"Description"]
        final_table.at[i,"Item Type"]=gerp_column.at[0,"Item Type"]
        final_table.at[i,"Qty Per Assembly"]=gerp_column.at[0,"Qty Per Assembly"]
        final_table.at[i,"Material Cost"]=gerp_column.at[0,"Material Cost"]
        final_table.at[i,"QPA*Material Cost"]=gerp_column.at[0,"QPA*Material Cost"]

        #match_column
        match_column=list(match_join.iloc[i].dropna(axis=0).drop(['Seq.','count','gerpSeq.']).index)
        #price comparison
        price_diff=round(final_table.at[i,"Material Cost (LOC)"]-final_table.at[i,"QPA*Material Cost"],8)
        if match_column==['gerp_true']: #sorting column 이 True 이면
                final_table.at[i,"match"]="True"
        elif match_column==['gerp_price']:
            if price_diff!=0:
                final_table.at[i,"match"]="False price"
            elif price_diff=='nan':
                final_table.at[i,"match"]="False price"
            else:
                final_table.at[i,"match"]="Substitute"
        
        elif match_column==['gerp_sub']: # exc, sub =>  같게 match에 표시
            if price_diff!=0:
                final_table.at[i,"match"]="Substitute False price"
            else:
                final_table.at[i,"match"]="Substitute"
        
        elif match_column==['gerp_parent']:
            if price_diff!=0:
                final_table.at[i,"match"]="False parent price"
            else:
                final_table.at[i,"match"]="False parent"

        elif match_column==['gerp_exc']: # exc, sub =>  같게 match에 표시
            if price_diff!=0:
                final_table.at[i,"match"]="Substitute False price"
            else:
                final_table.at[i,"match"]="Substitute"

        elif match_column==['gerp_re']:
            if price_diff!=0:
                final_table.at[i,"match"]="Substitute False price"
            else:
                final_table.at[i,"match"]="Substitute"

        else:
            print("no logic error")
        #나중에 채울 price match
        final_table.at[i,"price match"]=''


#MTL column
mtl=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/gerp.xlsx',sheet_name='MTL Cost')
mtl=mtl[['Item No','Item Cost','Material Overhead Cost','Creation Period']]
#mtl 부분 데이터 정리
mtl=mtl.rename(columns={'Item No':'Child Item','Creation Period':'PAC Creation','Item Cost':'MTL Cost','Material Overhead Cost':'MTL OH'})
mtl['MTL OH']=mtl['MTL OH'].round(8)
mtl['Net Material']=(mtl['MTL Cost']-mtl['MTL OH']).round(8)
mtl['mtl start']=''

#칼럼 순서 바꾸기
mtl=mtl[['mtl start','Child Item','PAC Creation','MTL Cost','Net Material','MTL OH']]

# vlookup 하고 내림차순 정리
mtl_final=pd.merge(final_table, mtl, on='Child Item', how='left')
final_table = mtl_final.sort_values(by=['Seq.'],ascending=True)

#price match column 채우기
final_table["price match"]=(final_table['Net Material']-final_table['Unit Price (USD)'])*final_table['Qty Per Assembly']
cast_to_type = {'price match': float} # round error -> datatype
final_table = final_table.astype(cast_to_type) # round error -> datatype
final_table["price match"]=final_table["price match"].round(8)


#### price calculate exception 1 -> FL
## 자신의  parent와 다른 child 매칭 
#Coil, Steel(STS) / Sheet,Steel(GI) (자신의 부모가 Cover, Rear)
for i in range(len(final_table)):
    parent_part=final_table.at[i,"Parent Part"] #gerp
    part_des=final_table.at[i,"Desc."] #npt
    if part_des=='Coil,Steel(STS)':
        for j in range(len(final_table)):
            another_child=final_table.at[j,"Child Item"]
            another_qty=final_table.at[j,"Qty Per Assembly"]
            if another_child==parent_part: # 한개 짜리 문제 될수도 있음
                string_condition=final_table.at[i,"Qty Per Assembly"]
                if type(string_condition)!=str:
                    final_table.at[i,"price match"]=float(final_table.at[i,"Net Material"]-final_table.at[i,"Unit Price (USD)"])*float(final_table.at[i,"Qty Per Assembly"])*float(another_qty)
    elif part_des=='Sheet,Steel(GI)':
        for j in range(len(final_table)):
            another_child=final_table.at[j,"Child Item"]
            another_des=final_table.at[j,"Desc."]
            another_qty=final_table.at[j,"Qty Per Assembly"]
            if another_child==parent_part and another_des=='Cover,Rear': # add one more condition
                final_table.at[i,"price match"]=(final_table.at[i,"Net Material"]-final_table.at[i,"Unit Price (USD)"])*final_table.at[i,"Qty Per Assembly"]*another_qty
    else:
        pass


#### price calculate exception 2 -> FL
## 자신의 child 다른 parent 매칭 but substitute part
#Tub,Drum(Front), Tub,Drum(Center), Tub,Drum(Rear)
for i in range(len(final_table)):
    child_part=str(final_table.at[i,"Child Item"])[:-2]  #gerp
    part_des=final_table.at[i,"Desc."] #npt
    part_match=final_table.at[i,"match"]
    if part_des.__contains__('Tub,Drum(') and part_match.__contains__('Substitute'):
        for j in range(len(final_table)):
            another_parent=final_table.at[j,"Parent Part"][:-2]#npt
            another_qty=final_table.at[j,"Qty Per Assembly"]
            another_price=final_table.at[j,"Material Cost (LOC)"]
            another_des=final_table.at[j,"Desc."]
            another_match=final_table.at[j,"match"]
            if another_des=='Coil,Steel(STS)' and another_parent==child_part and another_match.__contains__("Substitute")==False:
                final_table.at[i,"price match"]=(final_table.at[i,"Net Material"]-another_price)*final_table.at[i,"Qty Per Assembly"]
    elif part_des.__contains__('Cover,Rear') and part_match.__contains__('Substitute'):
        for j in range(len(final_table)):
            another_parent=final_table.at[j,"Parent Part"][:-2]
            another_qty=final_table.at[j,"Qty Per Assembly"]
            another_price=final_table.at[j,"Material Cost (LOC)"]
            another_des=final_table.at[j,"Desc."]
            another_match=final_table.at[j,"match"]
            if another_des=='Sheet,Steel(GI)' and another_parent==child_part and another_match.__contains__("Substitute")==False:
                final_table.at[i,"price match"]=(final_table.at[i,"Net Material"]-another_price)*final_table.at[i,"Qty Per Assembly"]
    else:
        pass

#### price calculate exception 3 -> FL (2/20)
## 자신의 child 다른 parent 매칭 but substitute part
#Tub,Drum(Front), Tub,Drum(Center), Tub,Drum(Rear)
for i in range(len(final_table)):
    gerp1_child=final_table.at[i,"Child Item"]
    gerp1_des=final_table.at[i,"Description"]
    gerp1_match=final_table.at[i,"match"]
    if gerp1_des.__contains__('Cover,Rear') and gerp1_match.__contains__('Substitute'):
        for j in range(len(final_table)):
            gerp2_parent=final_table.at[j,"Parent Item"]
            gerp2_des=final_table.at[j,"Description"]
            gerp1_cal=final_table.at[j,"Material Cost (LOC)"]
            if gerp2_des=='Sheet,Steel(GI)' and gerp2_parent==gerp1_child:
                final_table.at[i,"price match"]=(0-gerp1_cal)*final_table.at[i,"Qty Per Assembly"]
                final_table.at[j,"price match"]=np.nan
    else:
        pass

#price change -> True 
for i in range(len(final_table)):
    price_match=round(final_table.at[i,'price match'],2)
    final_match=final_table.at[i,"match"]
    if final_match!="False":
        if price_match==0 or price_match==0.00 or str(price_match)=='nan': 
            final_table.at[i,'match']=True
    else:
        final_table.at[i,"match"]=False

#substitute~~ ->Substitute, price~~~ -> Price Change
for i in range(len(final_table)):
    final_match=str(final_table.at[i,"match"])
    if final_match.__contains__('Substitute'):
        final_table.at[i,"match"]="Substitute"
    elif final_match.__contains__('price'):
        final_table.at[i,"match"]="Price Change"
                    
# Wihtin True, Price Change, part number not matching -> substitute 
for i in range(len(final_table)):
    final_match=str(final_table.at[i,"match"])
    npt_part=str(final_table.at[i,"Part No"])
    gerp_part=str(final_table.at[i,"Child Item"])
    if gerp_part!='':
        if npt_part!=gerp_part:
            final_table.at[i,"match"]="Substitute"


#price table -> final result
final_table['final result']=''
final_table.at[0,'Price Change']=0
final_table.at[0,'Substitute Price Change']=0
for i in range(len(final_table)):
    final_match=str(final_table.at[i,"match"])
    if final_match=='Substitute':
        price_value=str(final_table.at[i,"price match"])
        if price_value=='nan':
            pass
        else:
            final_table.at[0,'Substitute Price Change']=final_table.at[0,'Substitute Price Change']+final_table.at[i,"price match"]
    elif final_match=="Price Change":
        price_value=str(final_table.at[i,"price match"])
        if price_value=='nan':
            pass
        else:
            final_table.at[0,'Price Change']=final_table.at[0,'Price Change']+final_table.at[i,"price match"]

# #price match column round 2 -> data
final_table['Price Change']=round(final_table['Price Change'],2)
final_table['Substitute Price Change']=round(final_table['Substitute Price Change'],2)

#NPT에서 중복되는 항 제거 첫째항만 남기기 (일치와 substitute 경우)
for i in range(1,len(final_table)): #0은 -1과 비교할 수 없음으로
    npt_seq1=final_table.at[i,"Seq."]
    npt_seq2=final_table.at[i-1,"Seq."]
    if npt_seq1==npt_seq2:
        #sequence남기기
        final_table.at[i,"level"]=np.nan
        final_table.at[i,"Parent Part"]=np.nan
        final_table.at[i,"Part No"]=np.nan
        final_table.at[i,"Desc."]=np.nan
        final_table.at[i,"UOM"]=np.nan
        final_table.at[i,"Unit Qty"]=np.nan
        final_table.at[i,"Accum Qty"]=np.nan
        final_table.at[i,"Supplier Name"]=np.nan
        final_table.at[i,"Curr."]=np.nan
        final_table.at[i,"Unit Price"]=np.nan
        final_table.at[i,"Exchange Rate(LOC)"]=np.nan
        final_table.at[i,"Unit Price (USD)"]=np.nan
        final_table.at[i,"Material Cost (LOC)"]=np.nan


final_table.to_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/FL/final_table.xlsx')
