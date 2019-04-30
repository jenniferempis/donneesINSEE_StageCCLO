import os, xlrd, xlsxwriter
from xlrd import open_workbook
import pandas as pd

listeV = []

# Calcul année
annee=input("Entrez l'année de téléchargement : ")
year=int(annee)

# Répertoire
if os.path.exists(os.path.abspath('INSEE\\')+'\\')==True:
        folder = os.path.abspath('INSEE\\')+'\\'
else:
        os.mkdir(os.path.abspath('INSEE\\'))
        folder = os.path.abspath('INSEE\\')+'\\'

# Création d'un nouveau fichier
bdd = xlsxwriter.Workbook('INSEE\\A_integrer_'+annee+'.xlsx')
dpt64 = bdd.add_worksheet('a_integrer')

# Codes géos
code=input("Entrez l'année de téléchargement des codes géographiques : ")
cod=str(int(code)-2000)
wc = xlrd.open_workbook(folder+'table-appartenance-geo-communes-'+cod+'.xls',on_demand=True)
sheet0=wc.sheet_by_name("COM")

# Ajout des variables Codes géos

dpt64.write(0,0,sheet0.cell(5, 0).value)
listeV.append(sheet0.cell(5, 0).value)
dpt64.write(0,1,sheet0.cell(5, 1).value)
listeV.append(sheet0.cell(5, 1).value)
dpt64.write(0,2,sheet0.cell(5, 4).value)
listeV.append(sheet0.cell(5, 4).value)
dpt64.write(0,3,sheet0.cell(5, 7).value)
listeV.append(sheet0.cell(5, 7).value)
dpt64.write(0,4,sheet0.cell(5, 8).value)
listeV.append(sheet0.cell(5, 8).value)
dpt64.write(0,5,sheet0.cell(5, 9).value)
listeV.append(sheet0.cell(5, 9).value)
dpt64.write(0,6,sheet0.cell(5, 12).value)
listeV.append(sheet0.cell(5, 12).value)
dpt64.write(0,7,sheet0.cell(5, 15).value)
listeV.append(sheet0.cell(5, 15).value)

# Ajout des donnees Codes géos

constante2=1
for ligne in range(0,sheet0.nrows):
	if sheet0.cell_value(ligne, 2) == '64':
		dpt64.write(constante2,0,sheet0.cell(ligne, 0).value)
		dpt64.write(constante2,1,sheet0.cell(ligne, 1).value)
		dpt64.write(constante2,2,sheet0.cell(ligne, 4).value)
		dpt64.write(constante2,3,sheet0.cell(ligne, 7).value)
		dpt64.write(constante2,4,sheet0.cell(ligne, 8).value)
		dpt64.write(constante2,5,sheet0.cell(ligne, 9).value)
		dpt64.write(constante2,6,sheet0.cell(ligne, 12).value)
		dpt64.write(constante2,7,sheet0.cell(ligne, 15).value)
		constante2=constante2+1

# Recherche dans le dossier

lastCol=0
colDep = 0
firstRow = 0
ecart=0
nColF=8

dossier=str(year-3)+'_telechargement'+str(year)
folder_path=folder+dossier

for path, dirs, files in os.walk(folder_path):
        for filename in files:
                filename = os.path.join(path, filename)
                wb = xlrd.open_workbook(filename, '.xls',on_demand=True)
                if 'COM_'+str(year-3) in wb.sheet_names():
                        sheet1 = wb.sheet_by_name('COM_'+str(year-3))
                if 'ENSEMBLE' in wb.sheet_names():
                        sheet1 = wb.sheet_by_name('ENSEMBLE')
                if 'COM' in wb.sheet_names():
                        sheet1 = wb.sheet_by_name('COM')
  
                lastRow=sheet1.nrows
                lastCol=sheet1.ncols
                NumCol=[]
                constante=1
                # Recherche si DEP existe puis sauvergarde la ligne de début et la colonne de DEP
                for ligne in range(0,lastRow):
                        for col2 in range(0,lastCol):
                                if sheet1.cell(ligne, col2).value == 'DEP':
                                        colDep=col2
                                        firstRow=ligne
                                        break
                                if sheet1.cell(ligne, col2).value == 'CODGEO':
                                        colCod=col2
                                        firstRow=ligne
                                        break
                # Ajout des variables si elles n'existent pas
                for ligne in range(firstRow,lastRow):
                        for nCol in range(colDep+2,lastCol):
                                        if (sheet1.cell(ligne, colCod).value)[:2]=='64':
                                                if(sheet1.cell(firstRow, nCol).value in listeV)==False:
                                                        listeV.append(sheet1.cell(firstRow, nCol).value)
                                                        NumCol.append(nCol)
                                                        dpt64.write(0,nColF,sheet1.cell(firstRow, nCol).value)
                                                        nColF=nColF+1
                                                        break
                # Ajout des données correspondantes
                for ligne2 in range(0,lastRow):
                        if (sheet1.cell(ligne2, colCod).value)[:2]=='64':
                                nColF=nColF-len(NumCol)
                                for nCol2 in NumCol:
                                        dpt64.write(constante,nColF,sheet1.cell(ligne2, nCol2).value)
                                        nColF=nColF+1
                                constante=constante+1
                ecart=lastCol-len(NumCol)-colDep+2

bdd.close()

integration_bdd=pd.read_excel(folder+'A_integrer_'+annee+'.xlsx')
integration_bdd.to_csv(folder+'A_integrer_'+annee+'.csv',sep=';',index=False)

from sqlalchemy import create_engine

# Intégration du fichier en Base De Données
engine = create_engine('postgresql://postgres@localhost:5432/test')
nom_table = "bdd_"+annee
integration_bdd.to_sql(nom_table, engine,index=True, index_label='ogc_fid', if_exists='replace')

# Commentaires
Dico = xlrd.open_workbook(folder+'Dictionnaire_'+annee+'.xlsx',on_demand=True)
sheetD=Dico.sheet_by_index(0)
lastRow=sheetD.nrows
BDD= xlrd.open_workbook(folder+'A_integrer_'+annee+'.xlsx',on_demand=True)
dpt64=BDD.sheet_by_index(0)
for ligne3 in range(0,lastRow):
        if sheetD.cell(ligne3+1,1).value==dpt64.cell(0,ligne3).value:
                sqlString = "ALTER TABLE "+nom_table+ " COMMENT ON COLUMN " + nom_table.sheetD.cell(ligne3+1,1).value +" IS "+ sheetD.cell(ligne3+1,3).value
                data = pd.read_sql(sqlString, engine)
