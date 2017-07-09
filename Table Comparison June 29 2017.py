# -*- coding: utf-8 -*-
"""
Created on Thu Jun 29 08:30:35 2017

@author: hn7569
"""

import pandas as pd
from tkinter import *
from tkinter import filedialog
import time

print("Greetings! This analytic is used to compare two datasets but can also be used to join tables together for other types of analysis.")

print("\nWe will be matching two tables that you determine. TABLE X and TABLE Y.")

print("\nYou will still need to clean the fields before matching them in order to use application.")

input("\nPress ENTER when ready to select the TABLE X file.")

window=Tk()
window.filename = filedialog.askopenfilename(filetypes= (("All Files", "*.*"),("CSV Files","*.csv")))
filename=window.filename
window.destroy()
print( filename )

type=filename[-4:]

if type == ".xls":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            tablea = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
elif type == "xlsx":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            tablea = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
else:
    tablea = pd.read_csv(filename)


origColLista = list(tablea)

tablea['Record_ID'] = range(1, 1+len(tablea))

print("\nNext we will select what fields to match on. Please choose the same fields in the same order for both tables.")

print("\nColumns: ", origColLista)

print("\nWhat fields do you want to match from TABLE X? Enter 'd' when done.")

matcha=[]

while True:
    while True:
        fieldName = input("\nPlease enter column name (enter d if done): ")
        
        if fieldName == 'd' or fieldName in origColLista: break
        
        print("\nInvalid column name")
         
    if fieldName == 'd': break
    matcha.append(fieldName)
    
input("\nPress ENTER when ready to select the TABLE Y file.")

window=Tk()
window.filename = filedialog.askopenfilename(filetypes= (("All Files", "*.*"),("CSV Files","*.csv")))
filename=window.filename
window.destroy()
print( filename )

type=filename[-4:]

if type == ".xls":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            tableb = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
elif type == "xlsx":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            tableb = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
else:
    tableb = pd.read_csv(filename)

origColListb = list(tableb)

tableb['Record_ID'] = range(1, 1+len(tableb))

input("\nNext we will select what fields to match on for TABLE X. Please choose the same fields in the same order for both tables. Press ENTER.")

print("\nColumns: ", origColListb)

print("\nWhat fields do you want to match from TABLE Y? Enter 'd' when done.")

matchb=[]

while True:
    while True:
        fieldName = input("\nPlease enter column name (enter d if done): ")
        
        if fieldName == 'd' or fieldName in origColListb: break
        
        print("Invalid column name")
         
    if fieldName == 'd': break
    matchb.append(fieldName)
    
print("\nPlease wait while the the matching occurs...")

matched = tablea.merge(tableb, left_on=matcha, right_on=matchb, how='inner')

moa = tablea.merge(tableb, left_on=matcha, right_on=matchb, how='outer')

moa=moa[moa['Record_ID_x'].isnull()]

mob = tableb.merge(tablea, left_on=matchb, right_on=matcha, how='outer')

mob=mob[mob['Record_ID_x'].isnull()]

timestr = time.strftime("%Y%m%d")

filename = ("/Matched Output " + timestr + ".xls")

input("\nPress ENTER to indicate where you want the file to be exported.")

window=Tk()
window.filepath = filedialog.askdirectory()
filepath=window.filepath
window.destroy()

print("\nPlease wait while I prepare the file...")

writer = pd.ExcelWriter(filepath + filename)

matched.to_excel(writer,'Matched',index=False)

moa.to_excel(writer,'Table X Only',index=False)

mob.to_excel(writer,'Table Y Only',index=False)

writer.save()

print("\nSuccess! Result file: " + filepath + filename)

input("\nPress ENTER")