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


def generateRandomStudent(topicList, secList, verList):
    data = []
    characters = string.ascii_letters
    strRandName = ''.join([random.choice(characters) for i in range(random.randint(4, 7))])+ ' ' + ''.join([random.choice(characters) for i in range(random.randint(5, 8))])
    data.append(strRandName.title())
    data.append(secList[int(random.random() * len(secList))])
    data.append(verList[int(random.random() * len(verList))])
    for _ in topicList:
        data.append('Y' if random.random() < 0.5 else 'N')
    return data

def printQArray(fileOut, probTitle, topicTitle, probQuestions, probLeadIn, isTitled, numSections, numVersionsPerSec):
    fileOut.write('\\ifboolexpr{ togl {showAll} or togl {show'+probTitle+'}}{%\n')
    probVers = list(range(len(probQuestions)))
    if isTitled:
        fileOut.write('\\IndivExamTitled{'+topicTitle+'}{'+probLeadIn + '}%\n')
    else:
        fileOut.write('\\IndivExamUntitled{'+topicTitle+'}{'+probLeadIn + '}%\n')
    for _ in range(numSections):
        random.shuffle(probVers)
        for versNum in range(numVersionsPerSec):
            fileOut.write('{'+probQuestions[probVers[versNum]] + '}%\n')
    fileOut.write('\n}{}%\n')

def printQMulti(fileOutArray, probQuestions, numVersionsPerSec):
    probVers = list(range(len(probQuestions)))
    for idx, f in enumerate(fileOutArray):
        if idx % numVersionsPerSec == 0:
            random.shuffle(probVers)
        f.write(probQuestions[probVers[idx%numVersionsPerSec]])


def modifiedExamGeneration(examFileName, numVersions, numVersionsPerSec, sectionNames, versionTitles):
 #### Default Variables

    dir_path = os.path.dirname(os.path.realpath(__file__))

    os.chdir(dir_path)

    triggerStringOpen = '%**'
    triggerStringClose = '**%'

    secList = [s.strip() for s in sectionNames.split(',')]
    if versionTitles.find(',') == -1 and numVersionsPerSec > 1:
        verList = [chr(ord('A')+k) for k in range(numVersionsPerSec)]
    else:
        verList = [s.strip() for s in versionTitles.split(',')]
    numSections = len(secList)
    numVersionsPerSec = len(verList)

    print(secList)
    print(verList)

    ## Read through file to see if valid

    numProblems = 0
    topicHeaders = []
    topicCount = []
    validExam = True if numVersionsPerSec <= numVersions else False
    correctImport1 = False
    correctImport2 = False
    probTopicCount = 0
    whyNotValid = []
    needTitled = False
    needUntitled = False

    with open(examFileName + '.tex') as fileIn:
        for line in fileIn:
            if '\\usepackage{etoolbox}' in line:
                correctImport1 = True
            if '\\usepackage{xstring}' in line:
                correctImport2 = True
            if triggerStringOpen in line or '\\end{questions}' in line:
                if probTopicCount != numVersions  and len(topicHeaders) > 0:
                    whyNotValid.append('Topic ' + topicHeaders[-1] + ' does not have the right number of problems. It has ' + str(probTopicCount) + ' instead of ' + str(numVersions) + '.')
                    validExam = False
                if triggerStringOpen in line:
                    topicStr = line[line.index(triggerStringOpen)+3: line.index(triggerStringClose)]
                    if topicStr in topicHeaders:
                        topicCount[topicHeaders.index(topicStr)] += 1
                    else:
                        topicHeaders.append(topicStr)
                        topicCount.append(1)
                probTopicCount = 0
            if '\\question' in line or '\\titledquestion' in line:
                if '\\question' in line:
                    needUntitled = True
                else:
                    needTitled = True
                numProblems += 1
                probTopicCount += 1
        if numProblems % numVersions != 0:
            whyNotValid.append('The number of problems is not a multiple of the number of versions.')
            validExam = False

    print('Topics: ' + str(topicHeaders) + '\n')
    print('Number of Problems: ' + str(numProblems) + '\n')

    if len(topicHeaders) > 0 and not correctImport1:
        whyNotValid.append('Import of etoolbox is not present.')

    if not correctImport2:
        whyNotValid.append('Import of xstring is not present.')

    if not validExam:
        strError = 'This exam is not valid because:\n'
        for strx in whyNotValid:
            strError +='  ' + strx + '\n'
        tkmb.askokcancel("Invalid Exam", strError)
        exit()

    noTopics = len(topicHeaders) == 0

    if noTopics:
        print('No Topic Headers are present. Creating Individualized TeX files for each Section and Version.')
    else:
        print('Creating Exam based on the topic headers: ' + str(topicHeaders) + '.')

    numArg = numSections * numVersionsPerSec
    inQuestion = False
    probQuestions = []
    currentQuestion = ''
    topicNumber = 0
    probLeadIn = ''
    inLeadIn = False

    # What to do if no topic headers

    if noTopics:

        fileOutArray = []
        # Make a file for each different version of the exam being created

        for secNum in secList:
            for verNum in verList:
                fileOutArray.append(open(examFileName + '_' + secNum + '_' + verNum + '.tex', 'w'))


        # Go through the file, copying over all lines until you are in a question. 

        fileIn = open(examFileName + '.tex')
        for line in fileIn:
            if '\\end{questions}' in line:
                if inQuestion:
                    probQuestions.append(currentQuestion)
                    currentQuestion = ''
                if len(probQuestions) == numVersions:
                    printQMulti(fileOutArray, probQuestions, numVersionsPerSec)
                for f in fileOutArray:
                    f.write(line)
                inQuestion = False
            elif '\\question' in line or '\\titledquestion' in line: 
                if inQuestion:
                    # end current question
                    probQuestions.append(currentQuestion)
                    currentQuestion = ''
                    print(probQuestions)               
                if len(probQuestions) == numVersions:
                    printQMulti(fileOutArray, probQuestions, numVersionsPerSec)
                    probQuestions = []
                    currentQuestion = ''
                currentQuestion += line
                inQuestion = True
            elif inQuestion:
                currentQuestion += line
            else:
                for f in fileOutArray:
                    f.write(line)


        # Once in a question, build the array of questions like before.
        # Once the number of questions reaches the number of versions, print to all of the versions at once.
        # Close out at the end. 

        for f in fileOutArray:
            f.close()

    # This part here all assumes that you have topic headers.
    else:
        key =["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
        val = ["O", "T", "H", "F", "V", "X", "S", "E", "N", "Z"]

        probTitles = topicHeaders.copy() # Gets the objective names of problems on the exam
        for ind in range(len(key)):
            probTitles = [s.upper().replace(key[ind], val[ind]) for s in probTitles]

        topicNumber = 0 #len(topicHeaders)
        topicSetCount = 0

        with open(examFileName + '_MOD.tex', 'w') as fileOut:
            fileIn = open(examFileName + '.tex')
            for line in fileIn:
                if '\\gradetable' in line:
                    #fileOut.write(line.replace('\\gradeTable', '\\modGradeTable'))
                    fileOut.write('\\modGradeTable')
                elif '\\documentclass' in line and (line.find('%') < 0 or line.find('%') > line.find('\\documentclass')):
                    fileOut.write(line)
                    fileOut.write('\n\n%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%\n% Commands below added by the Individual Exam Generation Script\n% Only edit if you know what you are doing.\n% Matt Charnley, 2025\n')
                    if not correctImport1:
                        fileOut.write('\\usepackage{etoolbox}\n')
                    if not correctImport2:
                        fileOut.write('\\usepackage{xstring}\n')
                    # Add in toggles if needed
                    for x in probTitles:
                        fileOut.write('\\newtoggle{show'+x+'}\n')
                    fileOut.write('\\newtoggle{showAll}\n')
                    if needTitled:
                        strTitledCommand = ''
                        strTitledCommand += '\\newcommand{\\IndivExamTitled}[' + str(numArg+2) + ']{\n#2\n'## START HERE
                        strTitledCommand += '\\IfStrEq{\\versionPattern}{Z}\n{\n'
                        for secCount in range(numSections):
                            for versCount in range(numVersionsPerSec):
                                strTitledCommand += '\\titledquestion{#1} #'+str(secCount*numVersionsPerSec + versCount + 3)
                        strTitledCommand += '}{}\n'
                        for secCount in range(numSections):
                            for versCount in range(numVersionsPerSec):
                                strTitledCommand += '\\IfStrEq{\\versionPattern}{'+str(chr(65 + secCount*numVersionsPerSec + versCount))
                                strTitledCommand += '}\n{\\titledquestion{#1} #'+str(secCount*numVersionsPerSec + versCount + 3)
                                strTitledCommand += '}{}\n'
                        strTitledCommand += '}\n'
                        fileOut.write(strTitledCommand)
                    if needUntitled:
                        strUntitledCommand = ''
                        strUntitledCommand += '\\newcommand{\\IndivExamUntitled}[' + str(numArg+2) + ']{\n#2\n'## START HERE
                        strUntitledCommand += '\\IfStrEq{\\versionPattern}{Z}\n{\n'
                        for secCount in range(numSections):
                            for versCount in range(numVersionsPerSec):
                                strUntitledCommand += '\\question #'+str(secCount*numVersionsPerSec + versCount + 3)
                        strUntitledCommand += '}{}\n'
                        for secCount in range(numSections):
                            for versCount in range(numVersionsPerSec):
                                strUntitledCommand += '\\IfStrEq{\\versionPattern}{'+str(chr(65 + secCount*numVersionsPerSec + versCount))
                                strUntitledCommand += '}\n{\\question #'+str(secCount*numVersionsPerSec + versCount + 3)
                                strUntitledCommand += '}{}\n'
                        strUntitledCommand += '}\n'          
                        fileOut.write(strUntitledCommand)
                    # Add in command to arrange problem versions

                    # Write modGradeTableCommand
                    strModTable = ''
                    strModTable += '\n\\makeatletter\n\\newcommand{\\modGradeTable}{\n'
                    strModTable += '\\def\\tbl@range{AllQs}%\n\\if@addpoints\n\\@ifundefined{exam@numpoints}%\n{\\ClassWarning{exam}%\n'
                    strModTable += '{%\nYou must run LaTeX again to produce the\ntable.\\MessageBreak\n}%\n\\fbox{Run \\LaTeX{} again to produce the table}%\n}%\n'
                    strModTable += '{\\@ifundefined{range@\\tbl@range @firstq}%\n{\\range@undefined}%\n{%\n\\@ifundefined{range@\\tbl@range @lastq}%\n'
                    strModTable += '{\\range@undefined}%\n{%\n\\edef\\tbl@firstq{\\csname range@\\tbl@range @firstq\\endcsname}%\n\\edef\\tbl@lastq{\\csname range@\\tbl@range @lastq\\endcsname}%\n'
                    strModTable += '% Do not print table if firstq after lastq, i.e., if there are no questions:\n\\ifnum \\tbl@firstq > \\tbl@lastq\\relax%\n{}\n'
                    strModTable += '\\else%\n\\ifnum \\numquestions < 10\n\\multicolumnpartialpointtable{1}{AllQs}[questions]\n\\else\n\\ifnum \\numquestions > 18\n'
                    strModTable += '\\multicolumnpartialpointtable{3}{AllQs}[questions]\n\\else\n\\multicolumnpartialpointtable{2}{AllQs}[questions]\n\\fi\n\\fi\n\\fi\n'
                    strModTable += '}\n}\n}\n\\fi\n}\n\\makeatother\n'

                    fileOut.write(strModTable)

                    fileOut.write('\\providecommand{\\versionPattern}{Z}\n')
                    fileOut.write('\\providecommand{\\secNum}{}\n')
                    fileOut.write('\\providecommand{\\versNum}{}\n')
                    fileOut.write('\\providecommand{\\stuName}{}\n')
                    fileOut.write('%\\toggletrue{showAll}')
                    fileOut.write('% Uncomment the line above to show all problems, and add values into the above parameters to test what the individual exam will look like.\n')

                    fileOut.write('\n% End of additional commands.\n%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%*%\n\n')

                # begin questions
                elif '\\begin{questions}' in line:
                    fileOut.write(line)
                    fileOut.write('\n\\begingradingrange{AllQs}\n')
                # end questions

                elif '\\end{questions}' in line:
                    if inQuestion:
                        probQuestions.append(currentQuestion)
                        currentQuestion = ''
                    if len(probQuestions) == numVersions:
                        printQArray(fileOut, probTitles[topicNumber], topicHeaders[topicNumber], probQuestions, probLeadIn, isTitled, numSections, numVersionsPerSec)
                    fileOut.write('\\endgradingrange{AllQs}\n\\qformat{\\hfill This page is left blank for additional work \\hfill}\n\\ifprintanswers\n\\else\n')
                    fileOut.write('\\question\n\\newpage\n\\question\n%\\newpage\n%\\question\n\\fi\n')
                    fileOut.write(line)
                    inQuestion = False
                    inLeadIn = False


                elif '\\question' in line or '\\titledquestion' in line:
                    isTitled = '\\titledquestion' in line
                    inLeadIn = False
                    if inQuestion:
                        # end current question
                        probQuestions.append(currentQuestion)
                        currentQuestion = ''
                        print(probQuestions)               

                    if isTitled:
                        currentQuestion += line[line.index('}')+1:-1]
                    else:
                        currentQuestion += line[line.index('\\question')+9:-1]
                        print(currentQuestion)
                    inQuestion = True
                        

                    # fileOut.write('\n% THERE IS A QUESTION HERE\n\n')
                    # fileOut.write(line)
                elif triggerStringOpen in line:
                    if inQuestion:
                        probQuestions.append(currentQuestion)
                        currentQuestion = ''
                    if len(probQuestions) == numVersions:
                        printQArray(fileOut, probTitles[topicNumber], topicHeaders[topicNumber], probQuestions, probLeadIn, isTitled, numSections, numVersionsPerSec)
                        probQuestions = []
                        currentQuestion = ''
                        topicSetCount += 1
                        if topicSetCount == topicCount[topicNumber]:
                            topicNumber += 1
                            topicSetCount = 0
                                
                    fileOut.write(line)
                    inLeadIn = True
                    inQuestion = False

                elif inLeadIn:
                    probLeadIn += line
                elif inQuestion:
                    currentQuestion += line
                else:
                    fileOut.write(line)

            fileIn.close()
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = 'Template Question Data'

        sheet['A1'] = 'StudentName'
        sheet['B1'] = 'Section'
        sheet['C1'] = 'Version'
        

        for idx in range(len(topicHeaders)):
            sheet[chr(ord('D') + idx) + '1'] = '' + topicHeaders[idx] + '_' + str(topicCount[idx])

        numRandStudent = 20
        for idx in range(numRandStudent):
            dataRow = generateRandomStudent(topicHeaders, secList, verList)
            sheet.append(dataRow)
        
        os.chdir(examFileName[:examFileName.rindex('/')])
        wb.save('TemplateDatafor'+examFileName[examFileName.rindex('/')+1:] + '.xlsx')
        tkmb.showinfo('Generation Complete', 'Modified exam and template spreadsheet made successfully.\n\nExam is ' + os.path(fileOut)+'.\n\nSpreadsheet is ' + os.path.join(examFileName[:examFileName.rindex('/')],'TemplateDatafor'+examFileName[examFileName.rindex('/')+1:] + '.xlsx') +'.')