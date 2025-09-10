import os
import subprocess
import time
import sys
import datetime
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
import tkinter.messagebox as tkmb
import tkinter.simpledialog as tksd
import random
import openpyxl
import string
import shutil
from pypdf import PdfWriter, PdfReader
import pandas as pd
import ocrmypdf
from Levenshtein import distance

num_titlePages = 0
num_workPages = 0
colTitles = []

def processStudentExam(studentData, currentPage, pdfW, pdfROrig):
    for _ in range(num_titlePages):
        pdfW.add_page(pdfROrig.pages[currentPage])
        currentPage = currentPage + 1
    for ind in range(3, len(colTitles)):
        numProbs = int(colTitles[ind][colTitles[ind].index('_')+1:])
        if studentData[ind] == "Y":
            for _ in range(numProbs):
                pdfW.add_page(pdfROrig.pages[currentPage])
                currentPage = currentPage + 1
        else:
            for _ in range(numProbs):
                pdfW.add_blank_page()
                            
    for _ in range(num_workPages):
        pdfW.add_page(pdfROrig.pages[currentPage])
        currentPage = currentPage + 1
    return currentPage


def scannedExamProcessing(scanDirs, studentDataFile, numPagesBF, numPagesAL, nameLineText):
    global num_titlePages
    global num_workPages
    global colTitles
    num_titlePages = numPagesBF
    num_workPages = numPagesAL

    dataFrame = pd.read_excel(studentDataFile)
    dataFrame = dataFrame.fillna('')

    colTitles = dataFrame.columns
    cols = list(colTitles)

    for scans_Path in scanDirs:
        file_list = [f for f in os.listdir(path=scans_Path) if ((f.endswith('.pdf') or f.endswith('.PDF')) and not (f.startswith('OCR') or f.startswith('Pro') or f.startswith('x_')))]

        print('\nProcessing '+str(len(file_list)) +' files in directory ' + scans_Path + '.\n')

        os.chdir(scans_Path)

        for fileName in file_list:
            pdfW = PdfWriter()
            ocrmypdf.ocr(fileName, 'OCR_'+fileName, skip_text = True, deskew = True)
            pdfR = PdfReader('OCR_'+fileName)
            pdfROrig = PdfReader(fileName)

            currentPage = 0
            lastPage = len(pdfR.pages)

            while currentPage < lastPage:
                minDistance = 1000
                bestName = "XX"
                stuName = "XX"
                stuFound = False
                print("Page " + str(currentPage+1) + " in " + fileName)
                pageText = pdfR.pages[currentPage].extract_text()
                nLoc = pageText.find(nameLineText)
                if nLoc == -1:
                    stuName = tksd.askstring("Name: not found in text.", pageText + '\nEnter the student''s name from this text, or XX to skip the page.')
                else:
                    print(pageText)
                    nLocEnd = pageText.find("\n", nLoc)
                    print(nLoc)
                    print(nLocEnd)
                    
                    # TODO - add better filtering here based on the exam maybe?
                    # As the person inputs the names, look for them on the page and determine what was before/after them. Try to use that to more accurately get names in the future.

                    # My person version
                    # if nLocEnd == -1:
                    #     nLocEnd = pageText.find("NetID:", nLoc)
                    #     print(nLocEnd)
                    # if nLocEnd - nLoc < len(nameLineText) + 2:
                    #     print("Looking Further")
                    #     frontBoxEndOne = pageText.find("fully supported to receive credit.")
                    #     fBendOne = pageText.find("\n", frontBoxEndOne)
                    #     frontBoxEndTwo = pageText.find("must be fully")
                    #     fBendTwo = pageText.find("\n", frontBoxEndTwo)
                    #     print(fBendOne)
                    #     print(fBendTwo)
                    #     fBend = max(fBendOne, fBendTwo)
                    #     nameEnd = pageText.find("\n", fBend + 4)
                    #     stuName = pageText[fBend+1:nameEnd]
                    # else:
                    #     stuName = pageText[nLoc+len(nameLineText):nLocEnd]

                    if nLocEnd == -1 or nLocEnd - nLoc < len(nameLineText) + 2:
                        stuName = tksd.askstring('Could not find end of name', pageText + '\nCould not determine student name from text. Please input the name from Page ' + str(currentPage+1) + " in " + fileName + ' or \"XX\" to skip the page.')
                    else:
                        stuName = pageText[nLoc+len(nameLineText):nLocEnd]
                stuName = " ".join(stuName.split())
                newName = stuName
                if stuName == "XX":
                    currentPage = currentPage + 1
                else: 
                    rows = dataFrame.iterrows()
                    for r in rows:
                        rData = list(r[1])
                        nameTable = rData[0]
                        curDist = distance(nameTable.lower(), stuName.lower())
                        if curDist < minDistance:
                            minDistance = curDist
                            bestName = nameTable
                            print(str(curDist) + ' ' + bestName)
                        if nameTable.lower() == stuName.lower():
                            stuFound = True
                            studentData = rData
                    while not stuFound:
                        rows = dataFrame.iterrows()
                        if minDistance < 3:
                            tkmb.showinfo('Student Name Approximation', 'Student '+ newName +' not found. Using '+ bestName + ' instead. Distance is ' + str(minDistance) + '.')
                            newName = bestName
                        else:
                            bestCorrect = tkmb.askyesno('Student Not Found', 'Student '+ newName +' not found. Did you mean '+ bestName + '?')
                            if bestCorrect:
                                newName = bestName
                            else: 
                                newName = tksd.askstring('Next Name', 'Student '+ newName +' not found. \nWhat name to try next?')
                        newName = " ".join(newName.split())
                        for r in rows:
                            rData = list(r[1])
                            nameTable = rData[0]
                            curDist = distance(nameTable.lower(), stuName.lower())
                            if curDist < minDistance:
                                minDistance = curDist
                                bestName = nameTable
                                print(str(curDist) + ' ' + bestName)
                            if nameTable.lower() == newName.lower():
                                stuFound = True
                                studentData = rData
                                
                    currentPage = processStudentExam(studentData, currentPage, pdfW, pdfROrig)


            pdfW.write('Pro_'+fileName)

        os.chdir('..')
