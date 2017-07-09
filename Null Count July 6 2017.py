# -*- coding: utf-8 -*-
"""
Created on Wed Jul  5 14:43:47 2017

@author: hn7569
"""

import pandas as pd
from tkinter import *
from tkinter import filedialog
import xlsxwriter
import time

print("Greetings! This is a program that checks for NULL Values. ")

input("\nPress ENTER when ready to select input file.")

window=Tk()
window.filename = filedialog.askopenfilename(filetypes= (("CSV Files", "*.csv"),("All Files","*.*")))
filename=window.filename
window.destroy()
print( filename )

type=filename[-4:]

if type == ".xls":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            basicDF = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
elif type == "xlsx":
    while True:
        try:
            sn=input("\nWhat sheet in the file?")
            basicDF = pd.read_excel(filename,sheetname=sn)
            break
        except:
            print("\nThat was not a valid sheet name. Try again.")
else:
    basicDF = pd.read_csv(filename)

origColList = list(basicDF)

#Null Analysis Section

#Counting Number of Items      
total_entries=basicDF.count()

#Find Number of null in each field.
null_results=basicDF.isnull().sum()

#Merge Count and Null Results
null_output = pd.concat([total_entries,null_results],axis=1)

null_output.columns = ['NonMissing','Nulls']

null_output['Total Records']=null_output['Nulls']+null_output['NonMissing']

null_output=null_output.drop('NonMissing',axis=1)

null_output

null_output['Percentage Missing']=round((null_output['Nulls']/null_output['Total Records']),2)

#Taking a record count from that table before cleaning it up.
total_records=null_output[:1]
total_records=total_records['Total Records']
total_records=total_records.reset_index(drop=False)
total_records=total_records.drop('index',axis=1)
total_records

#Finalizing Null output table
null_output=null_output.reset_index(drop=False)
null_output.columns=['Field','Nulls','Total Records','Percentage Missing']
null_output=null_output.drop('Total Records',axis=1)

#Say what file you are profiling
description= "Data Profiling for "+filename

#File prep

timestr = time.strftime("%Y%m%d")
doc_date=time.strftime("%B %d, %Y")

filename = ("/Data_Nulls " + timestr + ".xlsx")

input("\nPress ENTER to indicate where you want the file to be exported.")

window=Tk()
window.filepath = filedialog.askdirectory()
filepath=window.filepath
window.destroy()

print("\nPlease wait while I prepare the file.")

writer = pd.ExcelWriter(filepath + filename, engine='xlsxwriter')

null_output.to_excel(writer,sheet_name='Nulls',index=False,startrow=7)

total_records.to_excel(writer,sheet_name='Nulls',index=False,startrow=4)

#Adding formats
workbook  = writer.book
worksheet = writer.sheets['Nulls']
format2 = workbook.add_format({'num_format': '0%'})
worksheet.set_column('C:C', None, format2)
title_font=workbook.add_format({'bold': True, 'font_size': 18})

#Adding documentation
worksheet.write_string(0, 0, description,title_font)
worksheet.write_string(1, 0, doc_date)

writer.save()


print("\nSuccess! Result file: " + filepath + filename)

input("\nPress ENTER")