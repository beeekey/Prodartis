# coding: latin1
import os, sys, glob
import openpyxl
import datetime
import unicodedata
from kitchen.text.converters import getwriter, to_bytes, to_unicode
from kitchen.i18n import get_translation_object
now=datetime.datetime.now()  

reload(sys)  
sys.setdefaultencoding('iso-8859-1')

files = glob.glob("*.stl")
files.sort(key=os.path.getmtime)

print(files[0])

import openpyxl
wb = openpyxl.Workbook()
wb.get_sheet_names()

sheet = wb.get_active_sheet()
sheet.title = 'prodartis_'+str(now.year)+'_'+str(now.month)+'_'+str(now.day)

wb.get_sheet_names()
sheet = wb.get_active_sheet()


### insert the title-row
i = 1
#title_row = ['Anzahl', 'Bauteilname', 'Technologie', 'Material', 'Nachbearbeitung', 'Farbe']
title_row = ['Anzahl/Pieces', 'Teilename/Partname', 'Material', 'Nachbearbeitung/Rework', 'Farbe/Color']
for j, title_val in enumerate(title_row):
    sheet[chr(65 + j) + str(i)] = title_val



### insert the filenames
i += 2
for file in files:
    partName=file.split(".")
    #partName=partName.replace("ü","ue")
    partName[0]=unicodedata.normalize('NFKD', u''+partName[0]+'').encode('iso-8859-1', 'ignore')
    print(partName[0])
    sheet['B'+str(i)] = partName[0]
    sheet['C'+str(i)] = 'PA12 HF'
    sheet['D'+str(i)] = 'gleitgeschliffen / nicht lackiert'
    sheet['E'+str(i)] = 'schwarz/black'
    i+=1

wb.save('prodartis_'+str(now.year)+'_'+str(now.month)+'_'+str(now.day)+'.xlsx')