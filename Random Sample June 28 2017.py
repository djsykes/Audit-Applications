# -*- coding: utf-8 -*-
"""
Created on Mon Jun 26 08:58:08 2017

@author: hn7569
"""
import pandas as pd
from tkinter import *
from tkinter import filedialog
import numpy as np
import time

print("Greetings! Need a random sample? Let's make one. Please import a sturctured dataset.")

input("\nPress ENTER when ready to select input file")

print("Please wait if it is a large file")

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

basicDF

sample_items=input("\nHow many items do you want to sample?")

sample = basicDF.take(np.random.permutation(len(basicDF))[:int(sample_items)])

timestr = time.strftime("%Y%m%d")

filename = ("/Sample Output " + timestr + ".xls")

input("\nPress ENTER to indicate where you want the file to be exported.")

window=Tk()
window.filepath = filedialog.askdirectory()
filepath=window.filepath
window.destroy()

print("\nPlease wait while I prepare the file.")

writer = pd.ExcelWriter(filepath + filename)

sample.to_excel(writer,'Sample',index=False)

writer.save()

print("\nSuccess! Result file: " + filepath + filename)

input("\nPress ENTER")


