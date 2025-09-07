import os
import subprocess
import time
import sys
import datetime
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
import tkinter.messagebox as tkmb
import random
import openpyxl
import string
import shutil
from pypdf import PdfWriter, PdfReader
import pandas as pd


def individualExamGeneration(examFileName, makeIndiv, studentDataFile,numVersions,secList, verList, outputFolder, blankFullExam, makeSolutions, combinedPDFs, makeScanDirect, sectionSubfolders):
    dir_path = os.path.dirname(os.path.realpath(__file__))

    os.chdir(dir_path)

    # Get Exam Parameters (get documentclass command)

    headerText = ''
    hasDCLine = False
    doneWithHeader = False
    inmyHeader = False

    restOfFile = ''

    with open(examFileName + '.tex', 'r') as fileIn:
        for line in fileIn:
            if '\\documentclass' in line:
                headerText += line
                hasDCLine = True
            elif not inmyHeader and '%*%*%*%*%*%*%' in line:
                inmyHeader = True
            elif not doneWithHeader and '%*%*%*%*%*%*%' in line:
                doneWithHeader = True
            elif (not hasDCLine) or (not doneWithHeader):
                headerText += line
            else:
                restOfFile += line

    if not hasDCLine:
        print('\\documentclass line not found.')
        exit()


    
    if (not makeIndiv) or studentDataFile == '':
        # Just make one exam of each type. Need to know how many to set up the code.
        if len(secList) == 0:
            print('List of sections not provided. Cannot generate exams.')
            exit()
        if len(verList) == 0:
            print('List of versions not provided.')
            if numVersions > 0:
                verList = [chr(ord('A') + idx) for idx in range(numVersions)]
                print('Defaulting to ' + verList)
            else:
                print('No number of versions provided. Cannot generate exams.')
                exit()   
        
        for secNum in secList:
            for versNum in verList:
                outputName = examFileName[:-4]+'_'+secNum+'_'+versNum
                variablesString = ''
                variablesString += '\\renewcommand{\\stuName}{}\n'
                variablesString+= '\\renewcommand{\\secNum}{'+secNum +'}'
                verPatt = chr(ord('A') + verList.index(versNum) + secList.index(secNum)*len(verList))
                variablesString += '\\renewcommand{\\versNum}{'+versNum+'}\\renewcommand{\\versionPattern}{'+verPatt+'}\n'
                variablesString += '\\toggletrue{showAll}'
                runFile = open(examFileName[:-4] + '_RUN.tex', 'w')
                runFile.write(headerText)
                runFile.write(variablesString)
                runFile.write(restOfFile)
                runFile.close()
                subprocess.check_call(['pdflatex', '-jobname', outputName, examFileName[:-4] + '_RUN.tex'])
                subprocess.check_call(['pdflatex', '-jobname', outputName, examFileName[:-4] + '_RUN.tex'])

                shutil.move(outputName + '.pdf', os.path.join(dir_path, outputFolder, outputName + '.pdf'))



    elif studentDataFile == '':
        print('Student Data file not provided.')
        exit()
    else:
        dataFrame = pd.read_excel(studentDataFile+'.xlsx')
        dataFrame = dataFrame.fillna('')

        colTitles = dataFrame.columns
        cols = list(colTitles)

        lastCol = len(cols)
        
        key =["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
        val = ["O", "T", "H", "F", "V", "X", "S", "E", "N", "Z"]

        objTitles = cols[3:lastCol]
        objCount = [int(s[s.index('_')+1:]) for s in objTitles]
        probTitles = [s[:s.index('_')] for s in objTitles] # Gets the objective names of problems on the exam
        for ind in range(len(key)):
            probTitles = [s.upper().replace(key[ind], val[ind]) for s in probTitles]

        secList = dataFrame['Section'].drop_duplicates().to_list()
        verList = dataFrame['Version'].drop_duplicates().to_list()

        if combinedPDFs:
            writer_dict = {}
            for s in secList:
                writer_dict[s] = PdfWriter()
        
        if sectionSubfolders:
            for s in secList:
                if not os.path.exists(os.path.join(outputFolder, s)):
                    os.mkdir(os.path.join(outputFolder, s))

        for row in dataFrame.itertuples():
            countN = 0
            outputName = examFileName[:-4]+'_' + row.Section + '_'+row.StudentName+'_'+row.Version
            variablesString = ''
            variablesString += '\\renewcommand{\\stuName}{'+row.StudentName+'}\n'
            variablesString+= '\\renewcommand{\\secNum}{'+row.Section +'}'
            verPatt = chr(ord('A') + verList.index(row.Version) + secList.index(row.Section)*len(verList))
            variablesString += '\\renewcommand{\\versNum}{'+row.Version+'}\\renewcommand{\\versionPattern}{'+verPatt+'}\n'
            
            for obj in objTitles:
                if getattr(row, obj) == 'Y':
                    variablesString += '\\toggletrue{show'+probTitles[objTitles.index(obj)] + '}\n'
                else:
                    countN += 1


            runFile = open(examFileName[:-4] + '_RUN.tex', 'w')
            runFile.write(headerText)
            runFile.write(variablesString)
            runFile.write(restOfFile)
            runFile.close()
            subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
            subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
            if combinedPDFs:
                writer_dict[row.Section].append('xxTEMPFILExx.pdf')
                if countN % 2 == 1:
                    writer_dict[row.Section].add_blank_page()
            if sectionSubfolders:
                shutil.move('xxTEMPFILExx' + '.pdf', os.path.join(dir_path, outputFolder, row.Section, outputName + '.pdf'))    
            else:
                shutil.move('xxTEMPFILExx' + '.pdf', os.path.join(dir_path, outputFolder, outputName + '.pdf'))

        if combinedPDFs:
            for s in secList:
                writer_dict[s].write(os.path.join(dir_path, outputFolder, examFileName[:-4]+'_' + s + '_ALL.pdf'))
        print('Individualized versions')

    if blankFullExam:
        for s in secList:
            for v in verList:
                outputName = examFileName[:-4]+'_' + s + '_'+ v + '_BLANK'
                variablesString = ''
                variablesString += '\\renewcommand{\\stuName}{}\n'
                variablesString+= '\\renewcommand{\\secNum}{'+s +'}'
                verPatt = chr(ord('A') + verList.index(v) + secList.index(s)*len(verList))
                variablesString += '\\renewcommand{\\versNum}{'+v+'}\\renewcommand{\\versionPattern}{'+verPatt+'}\n'
                variablesString += "\\toggletrue{showAll}"


                runFile = open(examFileName[:-4] + '_RUN.tex', 'w')
                runFile.write(headerText)
                runFile.write(variablesString)
                runFile.write(restOfFile)
                runFile.close()
                subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
                subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])

                shutil.move('xxTEMPFILExx' + '.pdf', os.path.join(dir_path, outputFolder, outputName + '.pdf'))
    
    if makeSolutions:
        for s in secList:
            for v in verList:
                outputName = examFileName[:-4]+'_' + s + '_'+ v + '_SOLUTIONS'
                variablesString = ''
                variablesString += '\\renewcommand{\\stuName}{}\n'
                variablesString+= '\\renewcommand{\\secNum}{'+s +'}'
                verPatt = chr(ord('A') + verList.index(v) + secList.index(s)*len(verList))
                variablesString += '\\renewcommand{\\versNum}{'+v+'}\\renewcommand{\\versionPattern}{'+verPatt+'}\n'
                variablesString += "\\toggletrue{showAll}\n\n\\printanswers"


                runFile = open(examFileName[:-4] + '_RUN.tex', 'w')
                runFile.write(headerText)
                runFile.write(variablesString)
                runFile.write(restOfFile)
                runFile.close()
                subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
                subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
                shutil.move('xxTEMPFILExx' + '.pdf', os.path.join(dir_path, outputFolder, outputName + '.pdf'))

        outputName = examFileName[:-4]+'_ALL_SOLUTIONS'
        variablesString = ''
        variablesString += '\\renewcommand{\\stuName}{}\n'
        variablesString+= '\\renewcommand{\\secNum}{All}'
        variablesString += '\\renewcommand{\\versNum}{All}\\renewcommand{\\versionPattern}{Z}\n'
        variablesString += "\\toggletrue{showAll}\n\n\\printanswers"

        runFile = open(examFileName[:-4] + '_RUN.tex', 'w')
        runFile.write(headerText)
        runFile.write(variablesString)
        runFile.write(restOfFile)
        runFile.close()
        subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
        subprocess.check_call(['pdflatex', '-jobname', 'xxTEMPFILExx', examFileName[:-4] + '_RUN.tex'])
        shutil.move('xxTEMPFILExx' + '.pdf', os.path.join(dir_path, outputFolder, outputName + '.pdf'))

        if makeScanDirect:
            for s in secList:
                for v in verList:
                    if not os.path.exists(os.path.join(outputFolder, 'Scans_' + s + '_' + v)):
                        os.mkdir(os.path.join(outputFolder, 'Scans_' + s + '_' + v))

    toDelete = [f for f in os.listdir(path=dir_path) if (f.startswith('xxTEMPFILExx') or f.startswith(examFileName[:-4] + '_RUN'))]
    for f in toDelete:
        os.remove(f) 
