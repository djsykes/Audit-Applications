# -*- coding: utf-8 -*-
"""
Created on Sun Jun 25 16:07:41 2017

@author: jmsyk
"""

import pandas as pd
from tkinter import *
from tkinter import filedialog
import time

print("Greetings! This is a duplicate test analytic. First define what is a duplicate, then we will put together the results!")

input("\nPress ENTER when ready to select input file.")

window=Tk()
window.filename = filedialog.askopenfilename(filetypes= (("CSV Files", "*.csv"),("All Files","*.*")))
filename=window.filename
window.destroy()
print( filename )

basicDF = pd.read_csv(filename)

origColList = list(basicDF)

print("\nColumns: ", origColList)

print("\nYou will now be asked for columns in which to check for duplicates")

dupColList=[]

while True:
    while True:
        fieldName = input("\nPlease enter column name (enter d if done): ")
        
        if fieldName == 'd' or fieldName in origColList: break
        
        print("Invalid column name")
         
    if fieldName == 'd': break
    dupColList.append(fieldName)
    
basicDF["Count"]=1 # augment basicDF with column "count"
dupColDF = basicDF[ dupColList + ["Count"] ]

dupCntDF = dupColDF.groupby(dupColList,as_index=False).sum()

dupCntDF = dupCntDF[dupCntDF['Count']>1]

nRows = len(dupCntDF)

dupCntDF['DupItemID'] = range(1, nRows+1)

result = basicDF.merge(dupCntDF, left_on=dupColList, right_on=dupColList, how='inner')

result = result.drop(['Count_x','Count_y'], axis=1)

result = result.sort_values('DupItemID', ascending=1)

timestr = time.strftime("%Y%m%d")

filename = ("/Duplicate Output " + timestr + ".xls")

input("\nPress ENTER to indicate where you want the file to be exported.")

window=Tk()
window.filepath = filedialog.askdirectory()
filepath=window.filepath
window.destroy()

writer = pd.ExcelWriter(filepath + filename)

result.to_excel(writer,'Results',index=False)

writer.save()

print("\nSuccess! Result file: " + filepath + filename)

input("\nPress ENTER")