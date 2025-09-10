import os
import subprocess
import time
import sys
import datetime
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
import tkfilebrowser
import tkinter.messagebox as tkmb
import random
import openpyxl
import string
import shutil
from pypdf import PdfWriter, PdfReader
import pandas as pd
import modifiedTexMethods as MT
import individualExamMethods as IE
import scannedExamMethods as SE

scanDirs = []
indivStudentDataPath = ''
examTeXFilePath = ''
output_path = ''

def resetAllWindows():
    wd_Database.withdraw()
    wd_IndivEx.withdraw()
    wd_ProbGen.withdraw()
    wd_Scanned.withdraw()
    wd_TexMod.withdraw()
    root.deiconify()


def databaseReader():
    root.withdraw()
    wd_Database.deiconify()
    #print("Database Reader")
    wd_Database.protocol("WM_DELETE_WINDOW", resetAllWindows)

def texModifier():
    root.withdraw()
    wd_TexMod.deiconify()
    #print("TeX Exam Modifier")
    wd_TexMod.protocol("WM_DELETE_WINDOW", resetAllWindows)

def probSelection():
    root.withdraw()
    wd_ProbGen.deiconify()
    #print("Problem Selection Generator")
    wd_ProbGen.protocol("WM_DELETE_WINDOW", resetAllWindows)

def individualExams():
    root.withdraw()
    wd_IndivEx.deiconify()
    indivStudentBoxes()
    #print("Individual Exam Generation")
    wd_IndivEx.protocol("WM_DELETE_WINDOW", resetAllWindows)

def scannedExams():
    root.withdraw()
    wd_Scanned.deiconify()
    #print("Scanned Exam Processing")
    wd_Scanned.protocol("WM_DELETE_WINDOW", resetAllWindows)

def open_File_TM():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = tkfd.askopenfilename(title="Select a file", filetypes=[ ["TeX Files", "*.tex"], ["All Files", "*.*"]], initialdir=dir_path)

    if file_path:
        dir_path, file_name = os.path.split(file_path)
        try:
            with open(file_path,'r') as file:
                content = file.read()
                lbl_tm_fileSelect["text"] = file_name
                os.chdir(dir_path)
                bool_hasFile = True
                global examTeXFilePath
                examTeXFilePath = file_path
        except Exception as e:
            tkmb.showerror("Error", f"Could not read file {e}")

def open_Exam_IE():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = tkfd.askopenfilename(title="Select a file", filetypes=[ ["TeX Files", "*.tex"], ["All Files", "*.*"]], initialdir=dir_path)

    if file_path:
        dir_path, file_name = os.path.split(file_path)
        try:
            with open(file_path,'r') as file:
                content = file.read()
                lbl_ie_fileSelect["text"] = file_name
                os.chdir(dir_path)
                bool_hasFile = True
                global examTeXFilePath
                examTeXFilePath = file_path
        except Exception as e:
            tkmb.showerror("Error", f"Could not read file {e}")

def open_Exam_IE_STU():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = tkfd.askopenfilename(title="Select a file", filetypes=[ ["Excel Files", "*.xlsx"], ["All Files", "*.*"]], initialdir=dir_path)

    if file_path:
        dir_path, file_name = os.path.split(file_path)
        try:
            with open(file_path,'r', encoding='latin-1') as file:
                content = file.read()
                lbl_ie_stuListSelect["text"] = file_name
                os.chdir(dir_path)
                bool_hasFile = True
                global indivStudentDataPath
                indivStudentDataPath = file_path
            dF = pd.read_excel(file_path)
            ie_secNames.set(', '.join(dF['Section'].unique()))
            ie_versNames.set(', '.join(dF['Version'].unique()))
        except Exception as e:
            tkmb.showerror("Error", f"Could not read file {e}")

def open_Output_IE():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    global output_path
    output_path = tkfd.askdirectory(title="Select an output directory for exam files:", initialdir=dir_path)
    if output_path:
        _, lastFolder = os.path.split(output_path)
        lbl_ie_dirSelect["text"] = lastFolder
        lbl_ie_outputDir["text"] = output_path

def fileReqs_TM():
    tkmb.askokcancel("File Requirements", "The TeX file you upload here must have the following properties:\n\n - Inclues the string '%**xx**%' before each topic of question you want sorted. \n\n - Include imports for ")

def modifiedExamGeneration():

    ## Data pulled as inputs
    dir_path = os.path.dirname(os.path.realpath(__file__))

    os.chdir(dir_path)

    #examFileName = lbl_tm_fileSelect["text"]
    #examFileName = examFileName[:-4]
    examFileName = examTeXFilePath[:examTeXFilePath.rindex('.')]
    numVersions = int(tm_numProbsSelected.get())
    numVersionsPerSec = int(tm_numVerSelected.get())
    sectionNames = ent_tm_sections.get()
    versionTitles = ent_tm_versions.get()

    MT.modifiedExamGeneration(examFileName, numVersions, numVersionsPerSec, sectionNames, versionTitles)

   

def individualExamGeneration():

    dir_path = os.path.dirname(os.path.realpath(__file__))

    os.chdir(dir_path)

    # Data pulled in as inputs:

    #examFileName = lbl_ie_fileSelect["text"][:lbl_ie_fileSelect["text"].rindex('.')]
    examFileName = examTeXFilePath[:examTeXFilePath.rindex('.')]
    makeIndiv = bool_ie_students.get()
    #studentDataFile = lbl_ie_stuListSelect["text"][:lbl_ie_stuListSelect["text"].rindex('.')]
    studentDataFile = indivStudentDataPath
    numVersions = int(ie_numVersSelected.get())
    sectionNames = ent_ie_sectionNames.get()
    secList = [s.strip() for s in sectionNames.split(',')]
    versionTitles = ent_ie_versNames.get()
    verList = [s.strip() for s in versionTitles.split(',')]
    outputFolder = output_path

    blankFullExam = bool_ie_blankExam.get()
    makeSolutions = bool_ie_makeSolns.get()
    combinedPDFs = bool_ie_combPDFs.get()
    makeScanDirect = bool_ie_scanDirect.get()
    sectionSubfolders = bool_ie_secSubfol.get()

    IE.individualExamGeneration(examFileName, makeIndiv, studentDataFile,numVersions,secList, verList, outputFolder, blankFullExam, makeSolutions, combinedPDFs, makeScanDirect, sectionSubfolders)

def indivStudentBoxes():
    if bool_ie_students.get():
        lbl_ie_stuListSelect.config(state=tk.ACTIVE)
        btn_ie_stuListFile.config(state = tk.ACTIVE)
    else:
        lbl_ie_stuListSelect.config(state = tk.DISABLED)
        btn_ie_stuListFile.config(state = tk.DISABLED)


def open_Dirs_Scan():
    global scanDirs
    dir_path = os.path.dirname(os.path.realpath(__file__))
    scanDirs = tkfilebrowser.askopendirnames(title = "Choose Directories that contain Scans", initialdir=dir_path)
    if len(scanDirs) == 0:
        lbl_scan_dirList["text"] = "No Directories Selected"
    else:
        strText = ""
        for s in scanDirs:
            h, tail = os.path.split(s)
            _, prevTail = os.path.split(h)
            strText = strText + os.path.join(prevTail, tail) + '\n'
        lbl_scan_dirList["text"] = strText

def open_Stu_Scan():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = tkfd.askopenfilename(title="Select a file", filetypes=[ ["Excel Files", "*.xlsx"], ["All Files", "*.*"]], initialdir=dir_path)

    if file_path:
        dir_path, file_name = os.path.split(file_path)
        try:
            with open(file_path,'r', encoding='latin-1') as file:
                content = file.read()
                lbl_scan_stuDataFile["text"] = file_name
                lbl_ie_stuListDir["text"] = dir_path
                os.chdir(dir_path)
                bool_hasFile = True
                global indivStudentDataPath
                indivStudentDataPath = file_path
        except Exception as e:
            tkmb.showerror("Error", f"Could not read file {e}")


def scannedExamProcessing():
    if len(indivStudentDataPath) == 0:
        tkmb.showerror("Error", "No student data file provided.")
    elif len(scanDirs) == 0:
        tkmb.showerror("Error", "No scan directories provided.")
    else:
        studentDataFile = indivStudentDataPath
        numPagesBF = int(scan_numBeforeFirst.get())
        numPagesAL = int(scan_numAfterLast.get())
        nameLineText = ent_scan_nameText.get()
        

        SE.scannedExamProcessing(scanDirs, studentDataFile, numPagesBF, numPagesAL, nameLineText)






root = tk.Tk()
s = ttk.Style()
root.title("Exam Generation Home Screen")
frm_root = ttk.Frame(root, padding = 10)
frm_root.grid()

#ttk.Button(frm_root, text = "Database Reader", command = databaseReader, padding = 10).grid(column=0, row = 0, sticky="EW")
ttk.Button(frm_root, text = "TeX Exam Modifier", command = texModifier, padding = 10).grid(column=0, row = 0, sticky="EW")
#ttk.Button(frm_root, text = "Problem Selection Generator", command = probSelection, padding = 10).grid(column=0, row = 2, sticky="EW")
ttk.Button(frm_root, text = "Individual Exam Generation", command = individualExams, padding = 10).grid(column=0, row = 1, sticky="EW")
ttk.Button(frm_root, text = "Scanned Exam Processing", command = scannedExams, padding = 10).grid(column=0, row = 2, sticky="EW")


wd_Database = tk.Toplevel(root)
wd_Database.title("Database Reader")
frm_wd = ttk.Frame(wd_Database, padding = 10)
frm_wd.grid()
wd_Database.withdraw()

wd_TexMod = tk.Toplevel(root)
wd_TexMod.title("TeX Exam Modifier")
frm_tm = ttk.Frame(wd_TexMod, padding = 10)
frm_tm.grid()
wd_TexMod.withdraw()

btn_tm_viewReqs = ttk.Button(frm_tm, text="View File Requirements", command = fileReqs_TM, padding =10)
lbl_tm_fileSelect = ttk.Label(frm_tm, text="No File Selected", justify=tk.CENTER, padding = 10)
lbl_tm_fileDir = ttk.Label(frm_tm, text="", width=10)
btn_tm_openFile = ttk.Button(frm_tm, text = "Open File", command=open_File_TM, padding = 10)
lbl_tm_numProblems = ttk.Label(frm_tm, text = "How many version of each problem are written? \n(Must be the same for each problem)", justify=tk.LEFT, padding = 10)
tm_numOptions = [1, 2, 3, 4, 5, 6,7, 8, 9]
tm_numProbsSelected = tk.StringVar()
tm_numProbsSelected.set("1")
tm_comboBox_probs = ttk.Combobox(frm_tm, textvariable=tm_numProbsSelected, values=tm_numOptions, state = "readonly")
lbl_tm_sections = ttk.Label(frm_tm, text = "What are the titles of the different sections? \nEnter as a comma-separated list.", justify = tk.LEFT, padding=10)
ent_tm_sections = ttk.Entry(frm_tm)
lbl_tm_numVersions = ttk.Label(frm_tm, text = "How many version of the exam do you want for each section?\n(Must be less than the number of problem versions)", justify=tk.LEFT, padding = 10)
tm_numVerSelected = tk.StringVar()
tm_numVerSelected.set("1")
tm_comboBox_vers = ttk.Combobox(frm_tm, textvariable=tm_numVerSelected, values=tm_numOptions, state = "readonly")
lbl_tm_versions = ttk.Label(frm_tm, text = "What are the titles of the different versions? \nEnter as a comma-separated list.", justify = tk.LEFT, padding=10)
ent_tm_versions = ttk.Entry(frm_tm)
btn_tm_generate = ttk.Button(frm_tm, text = "Generate Modified Exam", command=modifiedExamGeneration, padding = 10)


btn_tm_viewReqs.grid(column=0, row=0, columnspan=2)
lbl_tm_fileSelect.grid(column=0, row=1)
btn_tm_openFile.grid(column=1, row=1)
lbl_tm_numProblems.grid(column=0, row=2)
tm_comboBox_probs.grid(column=1, row=2)
lbl_tm_sections.grid(column=0, row=3)
ent_tm_sections.grid(column=1, row=3)
lbl_tm_numVersions.grid(column=0, row=4)
tm_comboBox_vers.grid(column=1, row=4)
lbl_tm_versions.grid(column=0, row=5)
ent_tm_versions.grid(column=1, row=5)
btn_tm_generate.grid(column= 0, row = 6, columnspan=2)



wd_ProbGen = tk.Toplevel(root)
wd_ProbGen.title("Problem Selection Generator")
frm_pg = ttk.Frame(wd_ProbGen, padding = 10)
frm_pg.grid()
wd_ProbGen.withdraw()

wd_IndivEx = tk.Toplevel(root)
wd_IndivEx.title("Individual Exam Generator")
frm_indiv = ttk.Frame(wd_IndivEx, padding = 10)
frm_indiv.grid()
wd_IndivEx.withdraw()


lbl_ie_fileSelect = ttk.Label(frm_indiv, text="No File Selected", justify=tk.CENTER, padding = 10)
lbl_ie_fileDir = ttk.Label(frm_indiv, text="", width=10)
btn_ie_openFile = ttk.Button(frm_indiv, text = "Select Modified Exam File", command=open_Exam_IE, padding = 10)
bool_ie_students = tk.BooleanVar()
bool_ie_students.set(False)
chkbx_ie_students = tk.Checkbutton(frm_indiv, text = "Make Individual Student Exams?", variable=bool_ie_students, command=indivStudentBoxes)
lbl_ie_stuListSelect = ttk.Label(frm_indiv, text="No File Selected", justify=tk.CENTER, padding = 10)
btn_ie_stuListFile = ttk.Button(frm_indiv, text = "Select Student Data List", command=open_Exam_IE_STU, padding = 10)
lbl_ie_stuListDir = ttk.Label(frm_indiv, text="", width=10)
ie_numOptions = [1, 2, 3, 4, 5, 6]
ie_numVersSelected = tk.StringVar()
ie_numVersSelected.set("2")
ie_comboBox_Vers = ttk.Combobox(frm_indiv, textvariable=ie_numVersSelected, values=ie_numOptions, state = "readonly")
lbl_ie_vers = ttk.Label(frm_indiv, text = "How many versions of each problem are in the exam?", justify=tk.CENTER, padding = 10)
lbl_ie_sectionNames = ttk.Label(frm_indiv, text = "What are the names of each of the sections? Enter as a comma-separated list.", justify=tk.CENTER, padding = 10)
ie_secNames = tk.StringVar()
ie_secNames.set("")
ent_ie_sectionNames = ttk.Entry(frm_indiv, textvariable=ie_secNames)
lbl_ie_versNames = ttk.Label(frm_indiv, text = "What name do you want to use for the different versions of the exam? Enter as a comma-separated list.", justify=tk.CENTER, padding = 10)
ie_versNames = tk.StringVar()
ie_versNames.set("")
ent_ie_versNames = ttk.Entry(frm_indiv, textvariable=ie_versNames)
lbl_ie_dirSelect = ttk.Label(frm_indiv, text="No Output Directory Selected", justify=tk.CENTER, padding = 10)
btn_ie_outputDir = ttk.Button(frm_indiv, text = "Select Output Directory", command=open_Output_IE, padding = 10)
lbl_ie_outputDir = ttk.Label(frm_indiv, text="", width=10)
frm_ie_Options = ttk.Frame(frm_indiv, padding=10, relief='sunken')
frm_ie_Options.grid()
btn_ie_generate = ttk.Button(frm_indiv, text = "Generate Individual Exam", command=individualExamGeneration, padding = 10)

bool_ie_blankExam = tk.BooleanVar()
bool_ie_blankExam.set(False)
bool_ie_makeSolns = tk.BooleanVar()
bool_ie_makeSolns.set(False)
bool_ie_combPDFs = tk.BooleanVar()
bool_ie_combPDFs.set(False)
bool_ie_scanDirect = tk.BooleanVar()
bool_ie_scanDirect.set(False)
bool_ie_secSubfol = tk.BooleanVar()
bool_ie_secSubfol.set(False)

s.configure('LJ.TCheckbutton', justify = tk.W)

chkbx_ie_blankExam = ttk.Checkbutton(frm_ie_Options, text = "Make blank exam for each version with all questions?", variable = bool_ie_blankExam, style = 'LJ.TCheckbutton')
chkbx_ie_makeSolns = ttk.Checkbutton(frm_ie_Options, text = "Make solution key for each version with all questions?", variable = bool_ie_makeSolns, style = 'LJ.TCheckbutton')
chkbx_ie_combPDFs = ttk.Checkbutton(frm_ie_Options, text = "Make combined PDF of exams for each section?", variable = bool_ie_combPDFs, style = 'LJ.TCheckbutton')
chkbx_ie_scanDirect = ttk.Checkbutton(frm_ie_Options, text = "Make directories in which to store scanned exams?", variable = bool_ie_scanDirect, style = 'LJ.TCheckbutton')
chkbx_ie_secSubfol = ttk.Checkbutton(frm_ie_Options, text = "Put individual exams into different folders based on section?", variable = bool_ie_secSubfol, style = 'LJ.TCheckbutton')

lbl_ie_fileSelect.grid(row = 0, column = 0)
btn_ie_openFile.grid(row = 0, column = 1)
chkbx_ie_students.grid(row = 1, column = 0)
lbl_ie_stuListSelect.grid(row = 1, column = 1)
btn_ie_stuListFile.grid(row = 1, column = 2)
ie_comboBox_Vers.grid(row = 2, column = 1)
lbl_ie_vers.grid(row = 2, column = 0)
lbl_ie_sectionNames.grid(row = 3, column = 0)
ent_ie_sectionNames.grid(row = 3, column = 1)
lbl_ie_versNames.grid(row = 4, column = 0)
ent_ie_versNames.grid(row = 4, column = 1)
lbl_ie_dirSelect.grid(row = 5, column=0)
btn_ie_outputDir.grid(row = 5, column=2)
frm_ie_Options.grid(row = 6, column = 0, columnspan=3)
chkbx_ie_blankExam.grid(row=0, column=0)
chkbx_ie_makeSolns.grid(row=1, column=0)
chkbx_ie_combPDFs.grid(row=2, column=0)
chkbx_ie_scanDirect.grid(row=3, column=0)
chkbx_ie_secSubfol.grid(row=4, column=0)
btn_ie_generate.grid(row= 7, column= 1)

wd_Scanned = tk.Toplevel(root)
wd_Scanned.title("Scanned Exam Processing")
frm_scan = ttk.Frame(wd_Scanned, padding = 10)
frm_scan.grid()
wd_Scanned.withdraw()

lbl_scan_dirList = ttk.Label(frm_scan, text = "No Directories Selected", justify=tk.CENTER, padding = 10)
lbl_scan_stuDataFile = ttk.Label(frm_scan, text = "No Student Data Provided", justify = tk.CENTER, padding = 10)
scan_numOptions = [0, 1, 2, 3, 4, 5, 6]
scan_numBeforeFirst = tk.StringVar()
scan_numBeforeFirst.set("2")
scan_numAfterLast = tk.StringVar()
scan_numAfterLast.set("2")
dD_scan_BF = ttk.Combobox(frm_scan, textvariable=scan_numBeforeFirst, values=scan_numOptions, state = "readonly")
dD_scan_AL = ttk.Combobox(frm_scan, textvariable=scan_numAfterLast, values=scan_numOptions, state = "readonly")
lbl_scan_BF = ttk.Label(frm_scan, text="How many pages (including the title) are before the first problem?")
lbl_scan_AL = ttk.Label(frm_scan, text="How many pages are after the last problem?")
lbl_scan_nameLine = ttk.Label(frm_scan, text="What text appears on the page right before the student's name?")
scan_nameText = tk.StringVar()
scan_nameText.set("Name:")
ent_scan_nameText = ttk.Entry(frm_scan, textvariable=scan_nameText)
btn_scan_selectDirs = ttk.Button(frm_scan, text = "Select Directories that contain scans.", command = open_Dirs_Scan, padding = 10)
btn_scan_selectStu = ttk.Button(frm_scan, text = "Select Individualized Student Data File.", command = open_Stu_Scan, padding = 10)
btn_scan_run = ttk.Button(frm_scan, text = "Process the Scans!", command = scannedExamProcessing, padding = 10)

lbl_scan_dirList.grid(column=0, row=1, rowspan = 4)
lbl_scan_stuDataFile.grid(column = 0, row = 0)
btn_scan_selectStu.grid(column=1, row= 0)
btn_scan_selectDirs.grid(column=1, row = 1)
lbl_scan_BF.grid(column = 1, row = 2)
lbl_scan_AL.grid(column = 1, row = 3)
lbl_scan_nameLine.grid(column = 1, row = 4)
dD_scan_BF.grid(column = 2, row = 2)
dD_scan_AL.grid(column = 2, row = 3)
ent_scan_nameText.grid(column = 2, row = 4)
btn_scan_run.grid(column = 1, row = 5, columnspan=2)



root.mainloop()
