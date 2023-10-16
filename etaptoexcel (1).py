import etap.api
import json
import sys
import xml.etree.ElementTree as ET
from etap.api.other.etap_client import EtapClient
import pandas as pd
import numpy as np
import regex as re
from datetime import datetime
import streamlit as st

#getting data from etap api and returning the connection object
def getconn():
    baseAddress = "http://localhost:49684"
    e = etap.api.connect(baseAddress)
    pingResult = e.application.ping()
    print(pingResult)
    return e

#creating an xml file "apiextract.xml" by calling api function getxml in the directory where program is being run from
def createxml():
    e=getconn()
    enames=json.loads(e.projectdata.getelementnames())
    print(enames)
    print(type(e.projectdata.getxml()))
    f=open("apiextract.xml","w")
    f.write(e.projectdata.getxml())
    return None

#checking for id descrepancies
def iddisc():
    dn=pd.read_excel(sst.existsheet,sheet_name=None)
    elementids=[]
    curtdic={}
    for i in list(dn.keys())[4:]:
        for j in range(2,dn[i].shape[0]):
            if dn[i].iloc[j,1]!=dn[i].iloc[j,1]:
                continue
            if re.search('Example',dn[i].iloc[j,1]):
                curtdic[i]=j
    for i in list(dn.keys())[4:]:
        for j in range(2,dn[i].shape[0]):
            elementids.append(dn[i].iloc[j,1])
    A=set(elementidlst)
    B=set(elementids)
    col[1].markdown(f"Elements not present in the model are {B-A}")
    col[1].markdown(f"Elements not present in the excel are {A-B}")
    return None

#creating connections dictionary fromto and tofrom
def connectionsdic(root):
    root=root
    fromlist=[]
    tolist=[]
    k=-1
    while True:
        k+=1
        try:
            fromlist.append(root[4][k].attrib['FromID'])
            tolist.append(root[4][k].attrib['ToID'])
        except Exception as ex:
            break
    fromtodic=dict(zip(fromlist,tolist))
    tofromdic=dict(zip(tolist,fromlist))
    return fromtodic,tofromdic

#function to make special changes to medium voltage protection setting   
def mvprotrans(createddataframedic):
    createddataframedic=createddataframedic
    ctlist=[]
    print('@@@@',df1['(6) MV protection settings'].iloc[1,1])
    idlist=list(createddataframedic[df1['(6) MV protection settings'].iloc[1,1]])
    awayrelay=[]
    awayctlist=[]
    awaycblist=[]
    nearrelay=[]
    nearctlist=[]
    nearcblist=[]
    print(tofromdic)
    print(fromtodic)
    for k in idlist:
        if k==k and k not in ['na','NA','Na','nan']:
            if tofromdic[k] in idlist:
                awayrelay.append(k)
                awayctlist.append(tofromdic[tofromdic[k]])
                awaycblist.append(tofromdic[tofromdic[tofromdic[k]]])
            else:
                nearrelay.append(k)
                nearctlist.append(tofromdic[k])
                nearcblist.append(tofromdic[tofromdic[k]])
    orderrelaylist=nearrelay+awayrelay
    ctlist=nearctlist+awayctlist
    cblist=nearcblist+awaycblist

    dicnew['(6) MV protection settings']=dicnew['(6) MV protection settings']
    dicnew['(6) MV protection settings']['ID_cat']=pd.Categorical(dicnew['(6) MV protection settings']['Relay ID'],categories=orderrelaylist,ordered=True)
    dicnew['(6) MV protection settings'].sort_values('ID_cat',inplace=True)

    dicnear=dic['CXFORM'].copy(deep=True)
    print(dicnear.shape,len(nearctlist))
    dicnear['ID_cat'] = pd.Categorical(dicnear['ID'], categories=nearctlist, ordered=True)
    dicnear.sort_values('ID_cat',inplace=True)

    dicaway=dic['CXFORM'].copy(deep=True)
    print(dicaway.shape,len(awayctlist))
    dicaway['ID_cat'] = pd.Categorical(dicaway['ID'], categories=awayctlist, ordered=True)
    dicaway.sort_values('ID_cat',inplace=True)
    
    print('@@',dicnew['(6) MV protection settings'].shape)
    print(len(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)])))
    dicnew['(6) MV protection settings']["Phase CT primary rated current\nIprim\n[A]"]=np.array(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))
    print('@@@@@@@@@@@@@@;\n\n\n\n\n',list(dicnear['ID'])+list(dicaway['ID'][:len(awayctlist)]))
    print('@@@@@@@@@@@@@@;\n\n\n\n\n',orderrelaylist)
    dicnew['(6) MV protection settings']["Phase CT secondary rated current\nIsec\n[A]"]=np.array(list(dicnear['Sec'])+list(dicaway['Sec'][:len(awayctlist)]))
    dicnew['(6) MV protection settings']["Phase CT\nClass"]=np.array(list(dicnear['IECBurdenDesignation'])+list(dicaway['IECBurdenDesignation'][:len(awayctlist)]))
    dicnew['(6) MV protection settings']["Phase CT Rated burden\nPr\n[VA]"]=np.array(list(dicnear['IECBurden'])+list(dicaway['IECBurden'][:len(awayctlist)]))
    dicnew['(6) MV protection settings']["CT Type"]=np.array(list(dicnear['Type'])+list(dicaway['Type'][:len(awayctlist)]))
    print(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)])))
    print(dicnew['(6) MV protection settings']["Overcurrent - Stage 1\n3I> / 51P-1\nI/Ir"])
    dicnew['(6) MV protection settings']["Overcurrent - Stage 1\n3I> / 51P-1\n[A]"]=(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))).apply(lambda x:float(x))*(pd.Series(list(dicnew['(6) MV protection settings']["Overcurrent - Stage 1\n3I> / 51P-1\nI/Ir"]))).apply(lambda x:float(x))
    dicnew['(6) MV protection settings']["Overcurrent - Stage 2\n3I>> / 51P-2\n[A]"]=(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))).apply(lambda x:float(x))*(pd.Series(list(dicnew['(6) MV protection settings']["Overcurrent - Stage 2\n3I>> / 51P-2\nI/Ir"]))).apply(lambda x:float(x))
    dicnew['(6) MV protection settings']["Overcurrent - Stage 3\n3I>>> / 50P\n[A]"]=(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))).apply(lambda x:float(x))*(pd.Series(list(dicnew['(6) MV protection settings']["Overcurrent - Stage 3\n3I>>> / 50P\nI/Ir"]))).apply(lambda x:float(x))
    dicnew['(6) MV protection settings']["Earth fault - Stage 1\nI0>\n[A]"]=(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))).apply(lambda x:float(x))*(pd.Series(list(dicnew['(6) MV protection settings']["Earth fault - Stage 1\nI0>\nI/Ir"]))).apply(lambda x:float(x))
    dicnew['(6) MV protection settings']["Earth fault - Stage 2\nI0>>\n[A]"]=(pd.Series(list(dicnear['Prim'])+list(dicaway['Prim'][:len(awayctlist)]))).apply(lambda x:float(x))*(pd.Series(list(dicnew['(6) MV protection settings']["Earth fault - Stage 2\nI0>>\nI/Ir"]))).apply(lambda x:float(x))
    return None

#function to make special changes to medium voltage protection setting  
def transproptrans():
    dicvecp1key={'0':'YN','1':'D'}
    dicvecp2key={'0':'yn','1':'d'}
    dicvecp3key={'0':'0','30':'1','60':'2','90':'3','120':'4','150':'5','180':'6','-150':'7','-120':'8','-90':'9','-60':'10','-30':'11'}
    #creating vecgroup column in dataframe as it is combination of different properties
    vecgrouplst=[]
    for k in range(0,len(dicnew["(3) Transformer"].iloc[:,1])):
        keyforvec=[]
        for l in ['PrimConnectionButton','PrimGroundingType','SecConnectionButton','SecGroundingType','PhaseShiftPS']:
            keyforvec.append(dic["XFORM2W"].loc[k,l])
        vecgrouplst.append(dicvecp1key[keyforvec[0]]+dicvecp2key[keyforvec[2]]+dicvecp3key[keyforvec[4]])
    dicnew[j]['Vector group']=vecgrouplst
    #transforming the primkv:
    dicnew[j]["Rated primary\nVoltage\nU1\n[kV]"]=dicnew[j]["Rated primary\nVoltage\nU1\n[kV]"]#to be converted
    #transforming date and time
    zlst=[]
    for z in list(dicnew[j]["Date"]):
        zlst.append(datetime.fromtimestamp(int(z)).strftime('%d-%m-%y'))
    dicnew[j]["Date"]=zlst
    #computing neutral earthin resistor ohms from amper for prim and secondary using the formula
    zlst=[]
    for z in range(0,len(dicnew[j]["Secondary side\nneutral earthing resistor\nRE\n[Ω]"])):
        if float(list(dic["XFORM2W"]["SecGroundingAmpers1"])[z])==0:
            print('00000000000')
            zlst.append(0)
            continue
        zlst.append(float(list(dic["XFORM2W"]["SeckV"])[z])*1000/(float(list(dic["XFORM2W"]["SecGroundingAmpers1"])[z])*1.732))
    dicnew[j]["Secondary side\nneutral earthing resistor\nRE\n[Ω]"]=zlst
    zlst=[]
    for z in range(0,len(dicnew[j]["Primary side\nneutral earthing resistor\nRE\n[Ω]"])):
        if float(list(dic["XFORM2W"]["PrimGroundingAmpers1"])[z])==0:
            zlst.append(0)
            continue
        zlst.append(float(list(dic["XFORM2W"]["PrimkV"])[z])*1000/(float(list(dic["XFORM2W"]["PrimGroundingAmpers1"])[z])*1.732))
    dicnew[j]["Primary side\nneutral earthing resistor\nRE\n[Ω]"]=zlst

#medium voltage switch gear adding frequency in frequency column
def mvsgtrans(): 
    dicnew[j]["Rated frequency\nfr\n[Hz]"]=[50]*len(dicnew[j]['MV switchgear ID'])

#main function
if __name__=="__main__":
    st.set_page_config(page_title='ETAP query',layout='wide')
    col=st.columns([0.2,0.6,0.2])
    sst=st.session_state

    if 'existsheet' not in st.session_state: #creating the form in streamlit
        sst.existsheet=''
        sst.buttonmemory1=False
        sst.keysheet=''
        sst.outputsheet=''
    
    with col[1].form("sub_form"):
        sst.existsheet=st.text_input('Input address of existing excel')
        sst.keysheet=st.text_input('Input address of key sheet')
        sst.outputsheet=st.text_input('Input address of output sheet')
        sst.submit1=st.form_submit_button("Submit")
    if sst.submit1:
        st.markdown('Form submitted')
        sst.buttonmemory1=True
    if not sst.buttonmemory1:
        st.markdown('exit1')
        st.stop()

    createxml()

    #using element tree library to parse the xml file (getting root of the xml)
    tree = ET.parse('apiextract.xml')
    root=tree.getroot()
    print(root)

    #creating the data frame i.e., dictionary with keys as element types and values as dataframes of property values from model
    dic={}
    k=-1
    elementidlst=[]
    while(True):
        k=k+1
        try:
            try:
                #bifurcating mvsg from lvsg by creating a seperate dataframe for the given conditions
                print('@@@@@@@@@@@@@@@',root[2][k].tag) 
                if root[2][k].tag=='BUS' and float(root[2][k].attrib['RatingType'])==2 and float(root[2][k].attrib['NominalkV'])>1:
                    print(dic['MVSG'])
                    dic['MVSG']=pd.concat([dic['MVSG'],pd.DataFrame(root[2][k].attrib,index=[0])], ignore_index=True)

                elif root[2][k].tag=='BUS' and float(root[2][k].attrib['RatingType'])==2 and float(root[2][k].attrib['NominalkV'])<=1:
                    print(dic['LVSG'])
                    dic['LVSG']=pd.concat([dic['LVSG'],pd.DataFrame(root[2][k].attrib,index=[0])], ignore_index=True)

                else:
                    print(dic[root[2][k].tag])
                    dic[root[2][k].tag]=pd.concat([dic[root[2][k].tag],pd.DataFrame(root[2][k].attrib,index=[0])], ignore_index=True)#if the element type already exists we concatnate the new dataframe
                elementidlst.append(root[2][k].attrib['ID'])

            except Exception as e:
                print(e)
                if root[2][k].tag=='BUS' and float(root[2][k].attrib['RatingType'])==2 and float(root[2][k].attrib['NominalkV'])>1:
                    dic['MVSG']=pd.DataFrame(root[2][k].attrib,index=[0])

                if root[2][k].tag=='BUS' and float(root[2][k].attrib['RatingType'])==2 and float(root[2][k].attrib['NominalkV'])<=1:
                    dic['LVSG']=pd.DataFrame(root[2][k].attrib,index=[0])
                
                else:
                    dic[root[2][k].tag]=pd.DataFrame(root[2][k].attrib,index=[0])
                print(root[2][k].tag)
                elementidlst.append(root[2][k].attrib['ID'])
        except:
            break
    if k+1==len(elementidlst):
        print('element list created without error')

    iddisc()

    #creating an excel of all the elements and all the properties
    with pd.ExcelWriter('output2.xlsx') as writer:
        for k in list(dic.keys()):
            dic[k].to_excel(writer,sheet_name=k)

    #creating an excel based on the template given (here df is a dictionary)
    df=pd.read_excel(sst.keysheet,sheet_name=None)
    df1={}
    for i in list(df.keys()):
        j=i
        try:
            df1[j]=df[i]
        except Exception as e:
            df1['N\A']=df[i]
    #removing unwanted rows from dataframes would not be required to be done as the excel is supposed to be clean

    #creating dictionary to map element type
    dictype={}
    list1=list(df1.keys())
    list2=[('UTIL'),(''),('XFORM2W'),('MVSG'),('HVCB'),('UNIVERSALRELAY'),('CABLE'),(''),(''),(''),('LVSG'),('LVCB'),('CABLE'),(''),(''),('')]
    print(len(list1))
    print(len(list2))
    dictype=dict(zip(list1,list2))

    #creating dictionary with keys as types to map each type element properties. masterdicelementytpe has keys as type and values as dictionary for mapping
    masterdicelementtype={}
    dicnew={}
    for j in df1.keys():
        masterdicelementtype[j]=dict(zip(list(df1[j].iloc[1:,1]),list(df1[j].iloc[1:,2])))
    createddataframedic={}

    fromtodic,tofromdic=connectionsdic(root)

    # dicnew[dictype[j]] = containing stuff
    # we are having one df for each type
    # dataframe is populated using dicdataframe
    # dicdataframe[propertyname]=dic[dictype[j]][propertyname]
    # this dicdataframe need not be multiple for same type
    # of element and all same type elements take\\\ properties
    # directly from this. dic has keys as elementtype and values as dataframes of all properties
    #creating final output dictionary and final excel out of it
    for j in list(df1.keys()):
        lencdf=0
        for i in df1[j].iloc[1:,1]: ####cycling through element properties in each j i.e., element type
            try:
                createddataframedic[i]=dic[dictype[j]][masterdicelementtype[j][i]]
            except Exception as e:
                try:
                    print(e,i)
                    createddataframedic[i]=['na']*len(dic[dictype[j]]['ID']) #if property is not mapped in the key excel this will be taken up
                except Exception as ee:
                    print('@@@@',ee)
                    createddataframedic[i]=['na']*len(dic[dictype[list(df1.keys())[0]]]['ID']) #if type is not mapped
                continue
        dicnew[j]=pd.DataFrame(createddataframedic) 
        ####if j is transformer type make changes in the dataframe (mva to kva and vectorgroup)
        if j == '(3) Transformer':
            transproptrans()
        if j == '(4) MV switchgear':
            mvsgtrans()
        if j == '(6) MV protection settings':
            print('bypassed')
            mvprotrans(createddataframedic)
            

        createddataframedic={}

    #creating final excel
    with pd.ExcelWriter(sst.outputsheet) as writer:
        print(writer.engine)
        for k in dicnew.keys():
            if k=='':
                dicnew[k].to_excel(writer,sheet_name='null')
                continue
            dicnew[k].to_excel(writer,sheet_name=k)