# -*- coding: utf-8 -*-
"""
Created on Fri Aug 23 13:58:40 2024

@author: krystad

Husk å få med OpenpyXL: https://stackoverflow.com/questions/73852273/openpyxl-not-found-in-exe-file-made-with-pyinstaller 
Anaconda prompt: pyinstaller oppmListe2.py --onefile --hidden-import openpyxl.cell._writer
"""

from pandas import read_csv, ExcelFile, read_excel, ExcelWriter, DataFrame
from numpy import where, nan
import sys 
import csv
import easygui
from datetime import datetime


def formatDate(d):
    return datetime(int(d[:4]),int(d[5:7]),int(d[8:]))

def sDate(dt):
    return str(dt).split(" ")[0]


intro = easygui.buttonbox('Dette programmet lager oppmøteliste. Du trenger:\n\n1. En Excel-fil fra Mentimeter\n2. En csv-fil fra Grupper på Blackboard\n\n\n\n','Oppmøteliste',('Videre (standard)','Videre (tilpass brukernavn)','Mer info','Avbryt'))

tilp = intro=='Videre (tilpass brukernavn)'

if intro in (None, 'Avbryt'):
    sys.exit()
    
elif intro == 'Mer info':
    infosvar = easygui.buttonbox('Hente ned gruppeliste fra Blackboard (kun første gang eller hvis gruppelista er endret):\n\n1. Inne på emnet ditt på Blackboard, gå til Grupper. \n   Trykk Eksporter --> Kun gruppemedlemmer.\n2. Du får en epost "Masseeksport fullført". Lagre fila.\n\nOpprette avstemning i Mentimeter (kun første gang):\n\n1. Gå til Mentimeter www.mentimeter.com/auth/saml/ntnu\n2. Trykk New Menti --> Start from scratch --> Open Ended\n3. Øverst skriver du navn på presentasjonen,\n   f.eks. "Oppmøte Teksam 1FA 2024/25".\n3. Bytt ut «Ask your question here…» med "Oppmøte: Skriv ditt NTNU‐brukernavn".\n\nAvstemning:\n\n1. I timen viser du Menti‐presentasjonen. Skru på QR‐kode.\n   Bruk samme presentasjon hver gang.\n2. Etter at alle har skrevet seg inn, trykk Manage Results --> Reset results.\n   Mentimeter har lagret resultatene, selv om du ikke ser dem.\n3. Finn presentasjonen i Mentimeter og trykk\n   View Results --> Download --> Spreadsheet (XLSX).\n\nDet kan hende studenter skriver brukernavnet feil. I så fall kan du trykke "Tilpass brukernavn".','Info',('OK','Avbryt'))
    if infosvar == 'Avbryt':
        sys.exit()
    

mentifil = easygui.fileopenbox('Velg Excel-fil fra Mentimeter', 'Oppmøte', '*.xlsx')

if mentifil == None:
    sys.exit()
studentfil = easygui.fileopenbox('Velg csv-fil fra Blackboard','Studentliste', '*.csv')
if studentfil == None:
    sys.exit()
    
dfalle = read_csv(studentfil,names=("Klasse","Brukernavn","Nr","Fornavn","Etternavn"),sep=',\s*', engine='python')

def remSpace(s):
    if type(s)==str:
        return s.replace('"', '')

dfalle = dfalle.map(remSpace)
    

try:
    dfalle = dfalle.drop('Nr', axis=1)
except KeyError:
    pass

antArk = len(ExcelFile(mentifil).sheet_names)

sisteDato,forsteDato=None,None
notFound = ""
notFoundNames,notFoundDates,ignoreNames=[],[],[]
nyeBrnavn = {}

for i in range(1,antArk):
    dfIn = read_excel(mentifil,sheet_name=i)
    dato = formatDate(dfIn["Unnamed: 1"][0])
    dfalle[dato]=nan
    if i==1:
        forsteDato = dato
    if i==antArk-1:
        sisteDato = dato
    kortdato = dato.strftime("%d.%m")
    for brnavn in dfIn["Question 1"][7:]:
        brnavn = str(brnavn).lower().split("@")[0].replace(" ","")
        if brnavn in ignoreNames:
            continue
        if brnavn in nyeBrnavn:
            brnavn = nyeBrnavn[brnavn]
        
        radarray = where(dfalle["Brukernavn"]==brnavn)[0]
        if len(radarray)==0:
            notFound+=f"{brnavn} ({kortdato})\n"
            notFoundNames.append(brnavn)
            notFoundDates.append(kortdato)
            if tilp:
                nyttBrnavn = easygui.enterbox(f'Brukernavn {brnavn} finnes ikke. Riktig brukernavn: (Skriv i for å ignorere)','Tilpass brukernavn')
                if nyttBrnavn == "i":
                    ignoreNames.append(brnavn)
                    print(f"{brnavn} ignorert")
                    continue
                elif nyttBrnavn is None:
                    easygui.msgbox('Programmet avsluttes.')
                    sys.exit()
                nyeBrnavn.update({brnavn:nyttBrnavn})
                try:
                    radarray = where(dfalle["Brukernavn"]==nyttBrnavn)[0]
                    dfalle.loc[radarray[0],dato]=1
                except IndexError:
                    easygui.msgbox(f'{nyttBrnavn} finnes heller ikke')
        elif len(radarray)==1:
            dfalle.loc[radarray[0],dato]=1

if notFound:
    notFoundMsg = easygui.buttonbox(f'Disse brukernavnene finnes ikke:\n{notFound}','Brukernavn ikke funnet på klasselista',('OK','Avbryt'))
    if notFoundMsg == "Avbryt":
        easygui.msgbox('Programmet avsluttes.')
        sys.exit()
else:
    easygui.msgbox('Alle brukernavn funnet på klasselista.')
    

dfalle.loc[len(dfalle)] = dfalle.sum(axis=0, numeric_only=True)

dfalle.insert(loc=0,column="Ganger",value=nan)
dfalle["Ganger"]=dfalle.sum(axis=1, numeric_only=True)

mask = dfalle["Ganger"] == 0
dfklasse = dfalle[~mask].reset_index().drop("index",axis=1)

dfklasse.loc[len(dfklasse)-1,"Ganger"]=None
dfklasse = dfklasse.rename(columns={d:d.strftime("%d.%m") for d in dfklasse.columns.values if type(d)==datetime})
dfklasse = dfklasse.sort_values(by="Ganger", ascending=False)

xlsnavn = mentifil[:-5]+"_fra_"+sDate(forsteDato)+"_til_"+sDate(sisteDato)+".xlsx"

notFoundDf=DataFrame({"Brukernavn":notFoundNames,"Dato":notFoundDates})

try:
    with ExcelWriter(xlsnavn) as writer:
        dfklasse.to_excel(writer,sheet_name='Oppmøte',index=False)
        #worksheet = writer.sheets['Oppmøte']
        #(max_row, max_col) = dfklasse.shape
        #worksheet.conditional_format(0, 0, max_row, 0, {"type": "3_color_scale"})
        if notFoundNames:
            notFoundDf.to_excel(writer,sheet_name="Ikke funnet",index=False)
except PermissionError:
    easygui.msgbox('Kan ikke skrive til Excel-fila. Du må lukke den. Programmet avsluttes.','Oppmøteliste')
    sys.exit()

easygui.msgbox(f'Oppmøte lagret i fil\n\n {xlsnavn}','Oppmøteliste')