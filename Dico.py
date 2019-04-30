import os
import xlrd
import xlsxwriter
from xlrd import open_workbook
from datetime import datetime

# Calcul année
annee=input("Entrez l'année de téléchargement : ")
year=int(annee)

# Répertoire
if os.path.exists(os.path.abspath('INSEE\\')+'\\')==True:
        folder = os.path.abspath('INSEE\\')+'\\'
else:
        os.mkdir(os.path.abspath('INSEE\\'))
        folder = os.path.abspath('INSEE\\')+'\\'

# création 
dico = xlsxwriter.Workbook('INSEE\\Dictionnaire_'+annee+'.xlsx')
dictionnaire = dico.add_worksheet("Dictionnaire")
dictionnaire.write(0,0,'VAR_INTEGRE')
dictionnaire.write(0,1,'VAR_ID')
dictionnaire.write(0,2,'VAR_LIB')
dictionnaire.write(0,3,'VAR_LIB_LONG')
listeV = []

# Répertoire
folder = os.path.abspath('INSEE\\')+'\\'

# Codes géos
wc = xlrd.open_workbook(folder+'table-appartenance-geo-communes-'+str(year-2000)+'.xls',on_demand=True)
sheet0=wc.sheet_by_name("Variables")
lastRow=sheet0.nrows
lastCol=sheet0.ncols

# Ajout des variables Codes géos

for ligne in range(0,lastRow):
        for col in range(0,lastCol):
                if sheet0.cell(ligne, col).value == 'VAR_ID':
                        firstRow=ligne
                        colVar=col
dictionnaire.write(1,0,"GEO_"+sheet0.cell(firstRow+1, 0).value)
dictionnaire.write(2,0,"GEO_"+sheet0.cell(firstRow+2, 0).value)
dictionnaire.write(3,0,"GEO_"+sheet0.cell(firstRow+5, 0).value)
dictionnaire.write(4,0,"GEO_"+sheet0.cell(firstRow+8, 0).value)
dictionnaire.write(5,0,"GEO_"+sheet0.cell(firstRow+9, 0).value)
dictionnaire.write(6,0,"GEO_"+sheet0.cell(firstRow+10, 0).value)
dictionnaire.write(7,0,"GEO_"+sheet0.cell(firstRow+13, 0).value)
dictionnaire.write(8,0,"GEO_"+sheet0.cell(firstRow+16, 0).value)

for co in range(0,lastCol):
        dictionnaire.write(1,co+1,sheet0.cell(firstRow+1, co).value)
        listeV.append(sheet0.cell(firstRow+1, colVar).value)
        dictionnaire.write(2,co+1,sheet0.cell(firstRow+2, co).value)
        listeV.append(sheet0.cell(firstRow+2, colVar).value)
        dictionnaire.write(3,co+1,sheet0.cell(firstRow+5, co).value)
        listeV.append(sheet0.cell(firstRow+5, colVar).value)
        dictionnaire.write(4,co+1,sheet0.cell(firstRow+8, co).value)
        listeV.append(sheet0.cell(firstRow+8, colVar).value)
        dictionnaire.write(5,co+1,sheet0.cell(firstRow+9, co).value)
        listeV.append(sheet0.cell(firstRow+9, colVar).value)
        dictionnaire.write(6,co+1,sheet0.cell(firstRow+10, co).value)
        listeV.append(sheet0.cell(firstRow+10, colVar).value)
        dictionnaire.write(7,co+1,sheet0.cell(firstRow+13, co).value)
        listeV.append(sheet0.cell(firstRow+13, colVar).value)
        dictionnaire.write(8,co+1,sheet0.cell(firstRow+16, co).value)
        listeV.append(sheet0.cell(firstRow+16, colVar).value)

nLig=9
firstRow=0

dossier=str(year-3)+'_telechargement'+str(year)
folder_path=folder+dossier

for path, dirs, files in os.walk(folder_path):
        for filename in files:
                filenam = os.path.join(path, filename)
                wb = xlrd.open_workbook(filenam, '.xls',on_demand=True)
                SheetNameList = wb.sheet_names()
                
                if filename=='base-cc-caract-emploi-'+str(year-3)+'.xls':
                        nom='CARAC_EMP_'
                elif filename=='base-cc-coupl-fam-men-'+str(year-3)+'.xls':
                        nom='COU_FA_ME_'
                elif filename=='base-cc-diplomes-formation-'+str(year-3)+'.xls':
                        nom='DI_FO_'
                elif filename=='base-cc-emploi-pop-active-'+str(year-3)+'.xls':
                        nom='EMP_POP_A_'
                elif filename=='base-cc-evol-struct-pop-'+str(year-3)+'.xls':
                        nom='POP_'
                elif filename=='base-cc-logement-'+str(year-3)+'.xls':
                        nom='LOG_'
                elif filename=='FILO_DEC_COM':
                        nom='FILO_DEC_'

                if "Variables_"+str(year-3) in SheetNameList:
                        sheet = wb.sheet_by_name("Variables_"+str(year-3))
                        lastRow=sheet.nrows
                        lastCol=sheet.ncols
                        NumLig=[]
                        for ligne in range(0,lastRow):
                                for col in range(0,lastCol):
                                        if sheet.cell(ligne, col).value == 'DEP':
                                                firstRow=ligne
                                                break
                                        if sheet.cell(ligne, col).value == 'VAR_ID':
                                                colVar=col
                                                break
                        if firstRow!=0:
                                for ligne2 in range(firstRow+2,lastRow):
                                        if(sheet.cell(ligne2, colVar).value in listeV)==False:
                                                listeV.append(sheet.cell(ligne2, colVar).value)
                                                NumLig.append(ligne2)
                                                dictionnaire.write(nLig,0,nom+sheet.cell(ligne2, 0).value)
                                                for col2 in range(1,lastCol+1):
                                                        dictionnaire.write(nLig,col2,sheet.cell(ligne2, col2-1).value)
                                                nLig=nLig+1
                        ecart=lastRow-len(NumLig)-firstRow+2
                        nLig=nLig-ecart           
                        
dico.close() 


