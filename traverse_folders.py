#external libraries
from pathlib import Path
import fitz
from tkinter import messagebox

#internal modules
from extract_grades import *
from grades_to_array import *


transcript_folder = Path(r'''C:\Users\afrifa_k\Desktop\IB1 Reports''').glob('**/*')

subject_array = [[], [], [], []]
studentsemgrades = [[], [], [], [], []]
studentexamgrades = [[], [], [], [], []]
transcript_temp = r'''C:\Users\afrifa_k\Desktop\TEMPLATE 2019.xlsx'''

newstudents = []


def folder_to_temp(pathlist):
    pathlist = pathlist.glob('**/*.pdf')
    for path in pathlist:
        file = fitz.open(path)
        checkpage = file[0]
        text = checkpage.getText("text")
        text = text.split("\n")
        if text[1] == "A project of SOS-KINDERDORF INTERNATIONAL":
            if (text[7] == "SEMESTER: 1") and (text[8] == "YEAR: 2016/2017"):
                igstudent_subjects, ig1sem1_array, ig1exam1_array = ig_to_array(path)
                subject_array[0]  = igstudent_subjects
                studentsemgrades[0] = ig1sem1_array
                studentexamgrades[0] = ig1exam1_array
            if (text[7] == "SEMESTER: 2") and (text[8] == "YEAR: 2016/2017"):
                igstudent_subjects, ig1sem2_array, ig1exam2_array = ig_to_array(path)
                subject_array[1] = igstudent_subjects
                studentsemgrades[1] = ig1sem2_array
                studentexamgrades[1] = ig1exam2_array
            if (text[7] == "SEMESTER: 1") and (text[8] == "YEAR: 2017/2018"):
                igstudent_subjects, ig2sem1_array, ig2exam1_array = ig_to_array(path)
                subject_array[2] = igstudent_subjects
                studentsemgrades[2] = ig2sem1_array
                studentexamgrades[2] = ig2exam1_array
            if (text[7] == "SEMESTER: 2") and (text[8] == "YEAR: 2017/2018"):
                igstudent_subjects, ig2sem2_array, ig2exam2_array = ig_to_array(path)
                subject_array[3] = igstudent_subjects
                studentsemgrades[3] = ig2sem2_array
                studentexamgrades[3] = ig2exam2_array
        if text[6] == "DP1":
            dpstudent_subjects, ib1sem1_array, ib1exam1_array = dp_to_array(path)
            for i in range(len(dpstudent_subjects)):
                if dpstudent_subjects[i] == "Spanish ab intio SL":
                    dpstudent_subjects[i] = "Spanish ab initio SL"
            studentsemgrades[4] = ib1sem1_array
            studentexamgrades[4] = ib1exam1_array
        print(path)

    iggrades_to_temp(transcript_temp, subject_array, studentsemgrades, studentexamgrades)
    ibgrades_to_temp(path, dpstudent_subjects, studentsemgrades, studentexamgrades)


def ibfolder_to_temp(pathlist):
    pathlist = pathlist.glob('**/*.pdf')
    for path in pathlist:
        file = fitz.open(path)
        checkpage = file[0]
        text = checkpage.getText("text")
        text = text.split("\n")
        if text[6] == "DP1":
            dpstudent_subjects, ib1sem1_array, ib1exam1_array = dp_to_array(path)
            for i in range(len(dpstudent_subjects)):
                if dpstudent_subjects[i] == "Spanish ab intio SL":
                    dpstudent_subjects[i] = "Spanish ab initio SL"
            studentsemgrades[0] = ib1sem1_array
            studentexamgrades[0] = ib1exam1_array
    ibgrades_to_temp(path, dpstudent_subjects, studentsemgrades, studentexamgrades)


def TestFunction(transcript_folder):
    for pathlist in transcript_folder:
        extension = '.pdf'
        count = sum(1 for _ in Path(pathlist).rglob(f'*{extension}'))
        if count >= 5:
            folder_to_temp(pathlist)
        elif count > 0 and count <= 4:
            ibfolder_to_temp(pathlist)
            print(pathlist.stem)
            messagebox.showerror("Error", str(pathlist.stem) + " requires special attention.")
        elif count == 0:
            messagebox.showerror("Error", "There are no report PDFs in " + str(pathlist.stem) + "'s transcript folder")
            continue


