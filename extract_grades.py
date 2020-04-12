#Library which converts tabular data in PDF to csv file. Used in "old_ig_to_array" function.
import tabula
import pandas

#For manipulation of PDF data.
import fitz

from titlecase import titlecase

#Array storing IB subjects done in school.
ibSubjects = ["English A Literature SL", "English A Literature HL", "English A Language and Literature SL",
              "Amharic A: Literature HL", "French A: Language and Literature SL", "Swahili A: Literature SL",
              "Spanish ab intio SL", "English B HL", "English B SL", "French B HL", "French B SL", "English B HL",
              "English B SL", "History SL", "History HL", "Information Technology in a Global Society HL",
              "	Information Technology in a Global Society SL", "Social and Cultural Anthropology HL",
              "Social and Cultural Anthropology SL", "Computer Science HL", "Computer Science SL", "Physics HL",
              "Mathematics HL", "Economics HL", "Economics SL", "Swahili ab initio SL",  "Theory of Knowledge",
              "Mathematical Studies SL", "Chemistry HL", "Chemistry SL", "Biology HL", "Biology SL", "Geography HL",
              "Geography SL", "Visual Art HL", "Visual Art SL", "Spanish ab initio SL", "Spanish ab Initio SL"
              ,"Mathematical Studies"]

#Array storing MYP subjects done in school.
mypSubjects = ["Language and literature: English Language and Literature", "Language acquisition: French Language", "Individuals and societies: Economics",
               "Individuals and societies: Geography", "Sciences: Chemistry", "Sciences: Physics", "Mathematics: Extended Mathematics",
               "Mathematics: Standard Mathematics", "Design: Design"]

igSubjects = ["BIOLOGY", "ENGLISH LITERATURE", "ECONOMICS", "MATHEMATICS", "FRENCH", "GEOGRAPHY", "ENGLISH LANGUAGE",
              "HISTORY", "CHEMISTRY", "COMPUTER SCIENCE", "PHYSICS", "ADDITIONAL MATHEMATICS", "ART AND DESIGN",
              "BUSINESS STUDIES"]

igSubjectsOld = []

for i in range(len(igSubjects)):
    igSubjectsOld.append(titlecase(igSubjects[i]))

# Function to store values collected from IB report in array. Much less contrived logic than "ig_to_array" Also much shorter.
def dp_to_array(filename):
    student_subjects = []
    ibsemgrades = []
    ibexamgrades = []
    doc = fitz.open(filename)
    # Make conditional to account for change in reports for IB1 second semester.
    page = doc[2]

    text = page.getText("text")

    student_data = text.split("\n")

    for i in range(0, len(student_data)):
        if student_data[9] == "Spanish ab Initio SL":
            student_data[9] = "Spanish ab initio SL"
        if student_data[i] in ibSubjects:
            student_subjects.append(student_data[i])
            if student_data[i + 1] == '7' or student_data[i + 1] == '6' or student_data[i + 1] == '5' or student_data[i + 1] == '4' or student_data[i + 1] == '3' or student_data[i + 1] == '2' or student_data[i + 1] == '1':
                ibexamgrades.append(student_data[i + 1])
                ibsemgrades.append(student_data[i + 2])
            else:
                ibsemgrades.append(student_data[i+3])
                ibexamgrades.append(student_data[i+2])
        if student_data[i] == "Aggregate":
            ibexamgrades.append(student_data[i+1])
    return student_subjects, ibsemgrades, ibexamgrades


def myp_to_array(filename):
    student_subjects = []
    mypgrades = []

    doc = fitz.open(filename)
    page1 = doc[6]
    page2 = doc[7]
    page3 = doc[8]
    text_page1 = page1.getText("text")
    text_page2 = page2.getText("text")
    text_page3 = page3.getText("text")
    all_text = text_page1 + text_page2 + text_page3
    student_data = all_text.split("\n")
    for i in range(0, len(student_data)):
        if student_data[i] in mypSubjects:
            student_subjects.append(student_data[i])
            if student_data[i + 2] == "Final Grade":
                mypgrades.append(student_data[i + 3])
            else:
                mypgrades.append(student_data[i + 2])
    return student_subjects, mypgrades

#Function to do same for new IG reports


def ig_to_array(filename):
    student_subjects = []
    igexamgrades = []
    igsemgrades = []

    doc = fitz.open(filename)
    checkpage = doc[0]
    text = checkpage.getText("text")
    student_data = text.split("\n")

    if student_data[1] == "A project of SOS-KINDERDORF INTERNATIONAL":
        page1 = doc[0]
        page2 = doc[1]
        text_page1 = page1.getText("text")
        text_page2 = page2.getText("text")
        all_text = text_page1 + text_page2
        student_data = all_text.split("\n")
        print(student_data)
        nationality = student_data[9]
        for i in range(len(student_data)):
            if student_data[i] in igSubjectsOld and (len(student_data[i + 1]) < 3):
                student_subjects.append(student_data[i])
                igsemgrades.append(student_data[i + 1])
                igexamgrades.append(student_data[i + 2])
                print(igexamgrades)
    else:
        page = doc[2]
        text = page.getText("text")
        student_data = text.split("\n")
    for i in range(0, len(student_data)):
        if student_data[i] in igSubjects:
            student_subjects.append(titlecase(student_data[i]))
            igsemgrades.append(student_data[i + 1])
            igexamgrades.append(student_data[i + 2])
    return student_subjects, igsemgrades, igexamgrades




