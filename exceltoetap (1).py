import etap.api
import json
import xml.etree.ElementTree as ET
from etap.api.other.etap_client import EtapClient
import pandas as pd
import numpy as np
import regex as re
import time
import streamlit as st

#creating element id descrepancy tracking
def iddisc():
    excelelementnamelist=[]
    curtdic={}
    for i in list(dn.keys())[4:]:
        for j in range(2,dn[i].shape[0]):
            if dn[i].iloc[j,1]!=dn[i].iloc[j,1]:
                continue
            if re.search('Example',dn[i].iloc[j,1]):
                curtdic[i]=j
    for i in list(dn.keys())[4:]:
        if i in list3:
            for j in range(2,dn[i].shape[0]):
                excelelementnamelist.append(dn[i].iloc[j,1])
    A=set(elementnamelist)
    B=set(excelelementnamelist)
    print(f"element ids not in excel {A-B}")
    print(f"element ids not in model {B-A}")
    return A-B,B-A

#transformer column transformations
def transcoltrans(i,k,dn,j):
    dn=dn
    j=j
    i=i
    dicvecp3key={'0': '0', '1': '30', '2': '60', '3': '90', '4': '120', '5': '150', '6': '180', '7': '-150', '8': '-120', '9': '-90', '10': '-60', '11': '-30'}
    vecgp=str(dn[i].iloc[j,k]).strip()
    if vecgp[0]=='D':
        root[2][dicxmlid[dn[i].iloc[j,1]]].set('PrimConnectionButton',str(1))
        root[2][dicxmlid[dn[i].iloc[j,1]]].set('PrimGroundingType',str(1))
        if vecgp[1]=='d':
           root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecConnectionButton',str(1))
           root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecGroundingType',str(1))
           root[2][dicxmlid[dn[i].iloc[j,1]]].set('PhaseShiftPS',str(dicvecp3key[vecgp[2]]))
        else:
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecConnectionButton',str(0))
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('PhaseShiftPS',str(dicvecp3key[vecgp[3:]]))
    else:
        try:
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('PrimConnectionButton',str(0))
        except Exception as E:
            print(E,'@@@@@@@@@@@@@@@@@@@',j,i,k)
            print(dn[i].iloc[j,1])
            print(dicxmlid[dn[i].iloc[j,1]])

        if vecgp[2]=='d':
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecConnectionButton',str(1))
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecGroundingType',str(1))
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('PhaseShiftPS',str(dicvecp3key[vecgp[3]]))
        else:
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('SecConnectionButton',str(0))
            root[2][dicxmlid[dn[i].iloc[j,1]]].set('PhaseShiftPS',str(dicvecp3key[vecgp[4]]))
    #converting primekv value in transformer (dividing by 1000)
    root[2][dicxmlid[dn[i].iloc[j,1]]].set('PrimKV',str(float(root[2][dicxmlid[dn[i].iloc[j,1]]].attrib['PrimkV'])/1000))
    # root[2][dicxmlid[dn[i].iloc[j,1]]].set('AlteredTime',str(time.mktime(time.strptime(root[2][dicxmlid[dn[i].iloc[j,1]]].attrib['AlteredTime'], '%d.%m.%Y'))))

#to convert normal bus to swtich gear, as the parameter that decides it is not given in the list
def switchgeartrans(dn,j,i):
    i=i
    dn=dn
    j=j
    root[2][dicxmlid[dn[i].iloc[j,1]]].set('RatingType',str('2'))#switch gear needs to be discussed

#creating 2 dictionaries for connections
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

#function for medium voltage protection elementtype specific changes to be made
def mvprotrans(dn,i,j,k,idlist,tofromdic):
    tofromdic=tofromdic
    idlist=idlist
    j=j
    ctcol=k
    awayrelay=[]
    awayctlist=[]
    awaycblist=[]
    nearrelay=[]
    nearctlist=[]
    nearcblist=[]
    for x in idlist:
        if dn[i].iloc[x,1]==dn[i].iloc[x,1] and dn[i].iloc[x,1] not in ['na','NA','Na','nan']:
            print('##@@#@#@',tofromdic)
            if tofromdic[dn[i].iloc[x,1]] in list(pd.Series(idlist).apply(lambda x:dn[i].iloc[x,1])):
                awayrelay.append(x)
                awayctlist.append(tofromdic[tofromdic[dn[i].iloc[x,1]]])
                awaycblist.append(tofromdic[tofromdic[tofromdic[dn[i].iloc[x,1]]]])
            else:
                nearrelay.append(x)
                nearctlist.append(tofromdic[dn[i].iloc[x,1]])
                nearcblist.append(tofromdic[tofromdic[dn[i].iloc[x,1]]])
    orderrelaylist=nearrelay+awayrelay
    ctlist=nearctlist+awayctlist
    cblist=nearcblist+awaycblist
    dicrelaytoct=dict(zip(orderrelaylist,ctlist))
    dicrelaytocb=dict(zip(orderrelaylist,cblist))
    for x in orderrelaylist:
        root[2][dicxmlid[dicrelaytoct[x]]].set('Prim',str(dn[i].iloc[x,k]))
        root[2][dicxmlid[dicrelaytoct[x]]].set('Sec',str(dn[i].iloc[x,int(k)+1]))
        root[2][dicxmlid[dicrelaytoct[x]]].set('IECBurdenDesignation',str(dn[i].iloc[x,int(k)+2]))
        root[2][dicxmlid[dicrelaytoct[x]]].set('IECBurden',str(dn[i].iloc[x,int(k)+3]))

#program starts from here
if __name__=="__main__":
    #creating streamlit page and initializing session state variables
    st.set_page_config(page_title='Create ETAP elements', layout='wide')
    col=st.columns([0.2,0.6,0.2])

    sst=st.session_state
    if "buttonmemory1" not in st.session_state:
        sst.buttonmemory1=False
        sst.inputsheet=''
        sst.keysheet=''

    if "buttonmemory2" not in st.session_state:
        sst.buttonmemory2=False
        sst.check2=False
        sst.cordic=''
        sst.submit2=False

    if "buttonmemory3" not in st.session_state:
        sst.buttonmemory3=False
        sst.check3=False

    #due to streamlit's logic and requirement to change the file below two blocks have come above
    if sst.check3:
        sst.buttonmemory3=True

    if sst.buttonmemory3:
        enew=etap.api.connect("http://localhost:49684")
        pingResult=enew.application.ping()
        print(pingResult)
        import base64
        fhand=open(r'output.txt')
        f=fhand.read()
        b = base64.b64encode(bytes(f, 'utf-8'))
        ghand=open(r'encodedextract.txt','wb') 
        ghand.write(b)

        #accessing the encoded text (we can simply use ghand.read)
        exhand=open(r'encodedextract.txt','r')
        ex=exhand.read()
        #passing api string to api
        st.markdown("@@@@@@@@@@@@@@")
        kk=enew.projectdata.sendpdexml(str(ex))
        print(kk)
        st.stop()

    #streamlit form creation
    with col[1].form("sub_form"):
        sst.keysheet=st.text_input('Inut the file address of key sheet')
        sst.inputsheet=st.text_input('Input the file address of the excel sheet containing values')
        sst.submit1=st.form_submit_button("Submit")
        st.markdown(sst.buttonmemory1)

    if sst.submit1:
        sst.buttonmemory1=True
        st.markdown('given')
    if not sst.buttonmemory1:
        col[1].markdown("Please enter all the values and submit the form")
        st.stop()

    #connecting to etap api
    baseAddress="http://localhost:49684"
    e=etap.api.connect(baseAddress)
    pingResult=e.application.ping()
    print(pingResult)

    #writing the extract file
    f=open('apiextract.xml',"w") 
    f.write(e.projectdata.getxml())

    #parsing the file for types of elements and their properties
    tree=ET.parse('apiextract.xml')
    root=tree.getroot()
    k=-1
    typelist=[]
    while(True):
        k=k+1
        try:
            typelist.append(root[2][k].tag)
        except Exception as err:
            print(err)
            break
    elementtypelist=list(pd.Series(typelist).unique().ravel())
    print(elementtypelist,'@@')

    #mapping element id names to serial number in the xml file to fetch them back
    dicxmlid={}
    for i in range(0,len(typelist)):
        dicxmlid[(root[2][i].attrib["ID"])]=i

    #making list of element names
    elementnamelist=[]
    for i in range(0,len(typelist)):
        elementnamelist.append(root[2][i].attrib['ID'])
    print(len(elementnamelist))

    # accessing excel to read element property from each excel page.
    df=pd.read_excel(sst.keysheet,sheet_name=None) 

    #mapping element type to actual type names
    list3=list(df.keys())
    list4=[('UTIL'),(''),('XFORM2W'),('BUS'),('HVCB'),('UNIVERSALRELAY'),('CABLE'),(''),('UNIVERSALRELAY'),(''),('BUS'),('LVCB'),('CABLE'),(''),(''),('')]#########################################################################
    diceletype=dict(zip(list3,list4))
    print(diceletype)

    #streamlit form for check by user
    with col[1].form("affirm_form"):
        st.markdown(diceletype)
        st.markdown('Is the above mapping correct')
        sst.check2=st.checkbox('Yes')
        sst.submit2=st.form_submit_button("Submit")

    if sst.check2:
        st.markdown("Mapping is correct")
        sst.buttonmemory2=True

    #taking input from user
    if not sst.buttonmemory2:
        col[1].markdown('Enter the corrected mapping in the same format as shown above')
        sst.cordic=col[1].text_input('')
        if sst.cordic:
            try:
                diceletype=eval(sst.cordic)
            except Exception as e:
                print(e)
                st.markdown('exit2')
                st.stop()
        else:
            st.markdown('exit3')
            st.stop()

    #mapping properties names
    diceleprop={}
    for i in list(df.keys()):
        list1=np.array(df[i].iloc[1:,1]).ravel()
        list2=np.array(df[i].iloc[1:,2]).ravel()
        diceleprop[i]=dict(zip(list1,list2)) #dictionary of dictionaries of elepropnames mapped
    print(diceleprop)

    fromtodic,tofromdic=connectionsdic(root)

    #inputting values to xml
    # giving properties to element using set funciton of xml file and the mappin dictionaries prepared
    dn=pd.read_excel(sst.inputsheet,sheet_name=None)
    print('@##@@$$@',list(dn.keys())[4:])
    for i in list(dn.keys())[4:]:
        if i in list3:   
            idlist=[]
            print('@@@@     @@@@@', elementnamelist,'\n',elementtypelist)
            if diceletype[i] in elementtypelist: #level i
                for j in range(2,dn[i].shape[0]):
                    if dn[i].iloc[j,1] in elementnamelist:#level j
                        print(dn[i].iloc[j,1])
                        for k in range(2,dn[i].shape[1]):
                            if dn[i].iloc[1,k]=='Vector group':
                                vecgrpcol=k
                            if dn[i].iloc[1,k]=="Phase CT primary rated current\nIprim\n[A]":
                                ctcol=k
                            print('@#@#@@##',dn[i].iloc[j,k],i,j,k)
                            if dn[i].iloc[j,k]==dn[i].iloc[j,k] and diceleprop[i][dn[i].iloc[1,k]]==diceleprop[i][dn[i].iloc[1,k]] and str(diceleprop[i][dn[i].iloc[1,k]]) not in ['nan','','na'] and str(dn[i].iloc[j,k]) not in ['nan','na','']:
                                try:
                                    print(i,j,k)
                                    print(diceletype[i],dn[i].iloc[j,1],diceleprop[i][dn[i].iloc[1,k]],str(dn[i].iloc[j,k]))
                                    print(root[2][dicxmlid[dn[i].iloc[j,1]]].attrib[diceleprop[i][dn[i].iloc[1,k]]])
                                    root[2][dicxmlid[dn[i].iloc[j,1]]].set(diceleprop[i][dn[i].iloc[1,k]],str(dn[i].iloc[j,k]))
                                    print('@@@@@@@',type(str(dn[i].iloc[j,k])),diceleprop[i][dn[i].iloc[1,k]])
                                    print('elemtype',diceletype[i],'\n','elementname',dn[i].iloc[j,1],'\n', 'elementprop', diceleprop[i][dn[i].iloc[1,k]],str(dn[i].iloc[j,k]),'prop set')
                                except Exception as err:
                                    print('@@',err)
                            else:
                                print("the values in prop don't exist or element prop names not mapped")
                        if i=='(3) Transformer':
                            print('@@@@@@@@######@@@@@@@',i)
                            transcoltrans(i,vecgrpcol,dn,j)
                        if i=='(4) MV switchgear':
                            switchgeartrans(dn,j,i)
                        if i=='(9) LV switchgear':
                            switchgeartrans(dn,j,i)
                        if i=='(6) MV protection settings':
                            print('##@#@#')
                            if j not in idlist:
                                idlist.append(j)
                        
                    else:
                        print(dn[i].iloc[j,1],'not in element name list')
                        continue
                if i=='(6) MV protection settings':
                    mvprotrans(dn,i,j,ctcol,idlist,tofromdic) 
            else:
                print(diceletype[i], 'not in element type list')
                continue
            print(dn.keys()) 
    tree.write("output.xml")
    tree.write("output.txt")

    #creating element id descrepancy tracking
    state1,state2=iddisc()

    #asking if user has created a new project
    with col[1].form("affirm_newproj"):
        st.markdown("Check the box and submit the form then create a new project and click on submit again")
        st.markdown(f"Not in excel {state1}")
        st.markdown(f"Not in model {state2}")
        sst.check3=st.checkbox('Yes')
        sst.submit3=st.form_submit_button("Submit")
        st.markdown(sst.submit3)



    if not sst.buttonmemory3:
        st.stop()

