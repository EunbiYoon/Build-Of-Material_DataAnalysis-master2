# if this code is not needed until 04/08/23, you can remove
#TL
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



#####################Coil,Steel => substitue part인 경우 /not top loader
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

#remain_match 존재 유무
remain_match_count=len(remain_match)
if remain_match_count==0:
    pass
else:
    #matching 된 행은 제거 & 재정렬
    A=remain_match["index"].tolist()
    remain_gerp=remain_gerp.drop(A,axis=0)
    remain_gerp.reset_index(inplace=True, drop=True)











###########
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