from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import datetime,platform,os,time,easygui

##########################################################################################################
def isLineNone(fline,fNoneCounter):
    fNoneCounter=0
    for i in fline:
        if i.value==None:
            fNoneCounter+=1
    return fNoneCounter==len(fline)
##########################################################################################################
def isPageEmpty(xlpage0):
    nonect=0
    if(xlpage0.max_row==1):
        return isLineNone(xlpage0[xlpage0.max_row],nonect)
    else:
        noneLines=0
        for i in xlpage0:
            nonect=0
            if(isLineNone(i,nonect)):noneLines+=1
        return noneLines==xlpage0.max_row
##########################################################################################################
def RemoveEmptyRows(xlpage):
    lineIndex=0
    lineToRemove=[]
    for line in xlpage:
        NoneCounter=0
        lineIndex+=1
        if isLineNone(line,NoneCounter):
            lineToRemove.append(lineIndex)
        else:
            if NoneCounter>0:
                return "RegInconsistencyError",lineIndex
    for i in reversed(lineToRemove):
        xlpage.delete_rows(i)
    return "RemotionSuccessful",0
##########################################################################################################
def isLeapYear(year):
    if(year%4==0):
        return (year%100)!=0
    else:
        return (year%400)==0
##########################################################################################################
def isDateValid(date):
    #date in format [d,m,y]
    NormalYear={1:30,2:28,3:30,4:31,5:30,6:31,7:30,8:30,9:31,10:30,11:31,12:30}
    LeapYear={1:30,2:29,3:30,4:31,5:30,6:31,7:30,8:30,9:31,10:30,11:31,12:30}
    if(date[1] not in range(1,13)):
        return False
    if(isLeapYear(date[2])):
        if(date[0]>LeapYear[date[1]]):
            return False
    else:
        if(date[0]>NormalYear[date[1]]):
            return False
    return True
##########################################################################################################
def AddUser(reg,row):
    #verify if the data that the user is trying to add is valid and without inconsistency
    if(len(row)!=3): return "InvalidRegError"
    for i in row:
        if(i==None): return "InvalidRegError"
    if(not(isDateValid(row[2]))):
        return "InvalidDateError"
    existingReg=[i[1] for i in reg['Usuarios']][1:]
    if(row[1] in existingReg): return "UserAlreadyCadError"
    date=datetime.date(row[2][2],row[2][1],row[2][0])
    reg['Usuarios'].append([row[0],row[1],date])
    #create the page of the user (which contains the user payments)
    newsheetname="Pg_User"+str(row[1])
    reg.create_sheet(newsheetname)
##########################################################################################################
def DelUser(reg,row):
    #verify if the data that the user is trying to add is valid and without inconsistency
    if(len(row)!=3): return "InvalidRegError"
    for i in row:
        if(i==None): return "InvalidRegError"
    if(not(isDateValid(row[2]))):
        return "InvalidDateError"
    date=datetime.date(row[2][2],row[2][1],row[2][0])
    if([row[0],row[1],date]==reg[row[1]]):
        reg['Usuarios'].delete_rows(row[1])
        #delete the page of the user (which contains the user payments)
        sheetToDelName="Pg_User"+str(row[1])
        reg.remove_sheet(sheet2del)
        return "RemotionSuccess"
    else:
        if([row[0],row[1],date]not in reg):
            return "UserNotRegistered"
        else:
            return "InconsistentDB_Error"
##########################################################################################################
def ClearNoneInWorkbook(wkbook2clPath):
    wkbook2cl=load_workbook(wkbook2clPath)
    for page in wkbook2cl:
        print(page)
        if(not(isPageEmpty(page))):
            BDFix,Mline=RemoveEmptyRows(page)
            print(BDFix)
    wkbook2cl.save(wkbook2clPath)
##########################################################################################################
def ClearScreen():
    if(platform.system()=="Windows"): os.system("cls")
    if(platform.system()=="Linux"): os.system("clear")
##########################################################################################################
def CheckHeaders(database):
            return (database['Usuarios']['A1']=="Nome" and database['Usuarios']['B1']=="Matricula" and database['Usuarios']['C1']=="DataMatricula")
##########################################################################################################
def InitBD(reg,filename):
    ct=0
    for page in reg:
        if(ct==0):
                if(page.title!="Usuarios"):
                    page.title="Usuarios"
                if(page.title=="Usuarios"):
                    if(not(CheckHeaders(reg))):
                        page['A1']="Nome"
                        page['B1']="Matricula"
                        page['C1']="DataMatricula"
                ct+=1
    ClearNoneInWorkbook(filename)
    reg.save(filename)
##########################################################################################################
def GetNewRegCode(bdname):
    useBd=load_workbook(bdname)
    return useBd['Usuarios'].max_row
##########################################################################################################
def FindCad(method,bd,toFind):
    Found=[]
    for line in bd['Usuarios']:
        if(method=='1'):
            if(toFind in line[0].value):
                Found.append([line[0],line[1],line[2]])
        if(method=='2'):
            if(line[1]==toFind):
                Found.append([line[0],line[1],line[2]])
        if(method=='3'):
            if(line[2]==toFind):
                Found.append([line[0],line[1],line[2]])
    return Found
##########################################################################################################
def isStrADate(dateTxt):#dd/mm/aaaa
    cnum=['0','1','2','3','4','5','6','7','8','9']
    fct=0
    for c in date:
        if(fct in [2,5]):
            if(c not in ['/','-']):
                fct+=1
                return False
        else:
            if(c not in cnum):
                fct+=1
                return False
    return True
##########################################################################################################
rel={"January":1,"February":2,"March":3,"April":4,"May":5,"June":6,"July":7,"August":8,"September":9,"October":10,"November":11,"December":12}
menuTxt="""
1)Adicionar novo cadastro
2)Remover um cadastro existente
3)Procurar cadastro
4)Registrar pagamento de dízimo
5)Gerar carteira de fiel
6)Sair
"""
searchOP="""
1)Filtrar por nome
2)Filtrar por matrícula
3)Filtrar por data de matrícula
"""
ValidOP=['1','2','3','4','5','6']
searchMethods=['1','2','3']
def Menu(bdfile,path):
    InitBD(bdfile,path)
    while(True):
        ClearScreen()
        op=0
        InitBD(bdfile,path)
        while(op not in ValidOP):
            ClearScreen()
            op=input(menuTxt)
        if(op=='1'):
            name=input("Nome:")
            y=int(time.strftime("%Y"))
            m=rel[time.strftime("%B")]
            d=int(time.strftime("%d"))
            code=GetNewRegCode(path)
            AddUser(bdfile,[name,code,[d,m,y]])
        if(op=='2'):
            findMethod=0
            while(findMethod not in searchMethods):
                ClearScreen()
                findMethod=input(searchOP)
            if(findMethod=='1'):
                target=input("Insira o nome do usuário a ser removido:")
            if(findMethod=='2'):
                target=input("Insira a matrícula do usuário a ser removido:")
            if(findMethod=='3'):
                target=input("Insira a data de matrícula do usuário a ser removido:")
            ToDel=FindCad(findMethod,bdfile,target)
            if(len(ToDel)>1):
                nf=str(len(ToDel))
                print("Foram encontrados mais de um usuário com o critério informado. Selecione o que deseja remover:\n")
                #print("{0:{space}}{1:20}{2:12}{3:10}".format("","Nome","Matrícula","Data de matrícula",space=(len(nf)+4)))
                print("-"*(42+len(nf)+1))
                ct=0
                for i in ToDel:
                    ct+=1
                    #print("{:{t}} - {:20}{:12}{:10}".format(ct,i[0],i[1],i[2],t=len(nf)+1))
                    print(str(ct).ljust(len(nf),' '),i[0].ljust(20,' '),i[1].ljust(12,' '),i[2].ljust(10,' '))
                    #ljust put the content more to left side of the reserved space
                    #rjust put it more to right side
                rgToDel=0
        if(op=='3'):
            findMethod=0
            while(findMethod not in searchMethods):
                ClearScreen()
                findMethod=input(searchOP)
            if(findMethod=='1'):
                target=input("Insira o nome a ser localizado:")
            if(findMethod=='2'):
                target=input("Insira a matrícula a ser localizada:")
            if(findMethod=='3'):
                target=input("Insira a data de matrícula a ser localizada:")
            res=FindCad(findMethod,bdfile,target)
            print("{:20}{:12}{:10}".format("Nome","Matrícula","Data de matrícula"))
            print("-"*42)
            for i in res:
                print("{:20}{:12}{:10}".format(i[0],i[1],i[2]))
        if(op=='4'):
            print("Nao implementada ainda")
            time.sleep(5)
        if(op=='5'):
            print("Nao implementada ainda")
            time.sleep(5)
        if(op=='6'):
            bdfile.save(path)
            break
file=easygui.fileopenbox("Selecione o arquivo XLSX a ser manipulado:",default="*.xlsx")
fileToUse=load_workbook(file)
Menu(fileToUse,file)
