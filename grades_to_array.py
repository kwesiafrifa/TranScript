import pandas as pd
from extract_grades import dp_to_array
from extract_grades import myp_to_array
from extract_grades import ig_to_array
from titlecase import titlecase
from transcript_formatting import format_page



def all_same(items):
    return all(x == items[0] for x in items)

def ibgrades_to_temp(path, dpstudent_subjects, studentsemgrades, studentexamgrades):
    aggregate1 = 0
    aggregate2 = 0
    for i in range(len(dpstudent_subjects)):
        if dpstudent_subjects[i] == "English A Language and Literature SL":
            dpstudent_subjects[i] = "English A Lang. & Lit. SL"
        if dpstudent_subjects[i] == "Mathematical Studies":
            dpstudent_subjects[i] = "Mathematical Studies SL"
        if dpstudent_subjects[i] == "Spanish ab Initio SL":
            dpstudent_subjects[i] = "Spanish ab initio SL"
    transcript = pd.read_excel("output.xlsx")
    for i in range (0, 23):
        if transcript.iloc[6, i] == "IB 1":
            if transcript.iloc[8, i] == "SG":
                for j in range(0, 38):
                    if transcript.iloc[j, 12] == "English A Lang. & Lit. SL" and (dpstudent_subjects[0] == "English A Literature HL" or "English A Literature SL"):
                        transcript.iloc[j, 12] = dpstudent_subjects[0]
                    if transcript.iloc[j, 12] == "History HL" and ("History SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "History SL"
                    if transcript.iloc[j, 12] == "Spanish ab initio SL" and ("Swahili ab initio SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Swahili ab initio SL"
                    if transcript.iloc[j, 12] == "French B HL" and ("French B SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "French B SL"
                    if transcript.iloc[j, 12] == "Economics HL" and ("Economics SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Economics SL"
                    if transcript.iloc[j, 12] == "Geography HL" and ("Geography SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Geography SL"
                    if transcript.iloc[j, 12] == "Computer Science HL" and ("Computer Science SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Computer Science SL"
                    if transcript.iloc[j, 12] == "ITGS HL" and ("Information Technology in a Global Society SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "ITGS SL"
                    if transcript.iloc[j, 12] == "ITGS HL" and ("Information Technology in a Global Society HL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "ITGS HL"
                    if transcript.iloc[j, 12] == "Soc. & Cult. Anthropology HL" and ("Social and Cultural Anthropology HL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Soc. & Cult. Anthropology HL"
                    if transcript.iloc[j, 12] == "Soc. & Cult. Anthropology HL" and ("Social and Cultural Anthropology SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Soc. & Cult. Anthropology SL"
                    if transcript.iloc[j, 12] == "Mathematics HL" and ("Mathematics SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Mathematics SL"
                    if transcript.iloc[j, 12] == "Biology HL" and ("Biology SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Biology SL"
                    if transcript.iloc[j, 12] == "Visual Art HL" and ("Visual Art SL" in dpstudent_subjects):
                        transcript.iloc[j, 12] = "Visual Art SL"

                    if transcript.iloc[j, 12] in dpstudent_subjects:
                        for k in range(len(dpstudent_subjects)):
                            if transcript.iloc[j, 12] == dpstudent_subjects[k]:
                                transcript.iloc[j, 13] = studentsemgrades[4][k]
                                transcript.iloc[j, 14] = studentexamgrades[4][k]
                                if studentexamgrades[4][k] != 'N/A':
                                    aggregate1 = aggregate1 + int(studentexamgrades[4][k])
                                    aggregate2 = aggregate2 + int(studentsemgrades[4][k])
                    if transcript.iloc[j, 12] == "AGGREGATE":
                        transcript.iloc[j, 13] = aggregate2
                        print(transcript.iloc[j, 13])
                        transcript.iloc[j, 14] = aggregate1


    transcript.to_excel(str(path.parent) + "/" + str(path.parent.stem) + "_" + "output.xlsx")
    format_page(str(path.parent) + "/" + str(path.parent.stem) + "_" + "output.xlsx", str(path.parent.stem))
    print(path.parent.stem)





def iggrades_to_temp(transcript_temp, subject_array, studentsemgrades, studentexamgrades):
 transcript_temp = pd.read_excel(transcript_temp)
 for i in range(0, 23):
        if transcript_temp.iloc[6, i] == "           IGCSE 1":
            if transcript_temp.iloc[8, i] == "SG":
                for j in range(0, 38):
                    for l in range(len(subject_array)):
                        if (transcript_temp.iloc[j, 0] in subject_array[l]) and (all_same(subject_array) is True):
                            for k in range(len(subject_array[l])):
                                if transcript_temp.iloc[j, 0] == subject_array[l][k]:
                                    transcript_temp.iloc[j, 1] = studentsemgrades[0][k]
                                    transcript_temp.iloc[j, 2] = studentexamgrades[0][k]
                                    transcript_temp.iloc[j, 3] = studentsemgrades[1][k]
                                    transcript_temp.iloc[j, 4] = studentexamgrades[1][k]
                                    transcript_temp.iloc[j, 5] = studentsemgrades[2][k]
                                    transcript_temp.iloc[j, 6] = studentexamgrades[2][k]
                                    transcript_temp.iloc[j, 7] = studentsemgrades[3][k]
                                    transcript_temp.iloc[j, 8] = studentexamgrades[3][k]
                        elif (transcript_temp.iloc[j, 0] in subject_array[l]) and (all_same(subject_array) is not True):
                            for k in range(len(subject_array[l])):
                                if transcript_temp.iloc[j, 0] == subject_array[0][k]:
                                    transcript_temp.iloc[j, 1] = studentsemgrades[0][k]
                                    transcript_temp.iloc[j, 2] = studentexamgrades[0][k]
                                    #Concept of storing dropped rows
                                if transcript_temp.iloc[j, 0] == subject_array[1][k]:
                                        transcript_temp.iloc[j, 3] = studentsemgrades[1][k]
                                        transcript_temp.iloc[j, 4] = studentexamgrades[1][k]
                                if transcript_temp.iloc[j, 0] == subject_array[2][k]:
                                        transcript_temp.iloc[j, 5] = studentsemgrades[2][k]
                                        transcript_temp.iloc[j, 6] = studentexamgrades[2][k]
                                if transcript_temp.iloc[j, 0] == subject_array[3][k]:
                                        transcript_temp.iloc[j, 7] = studentsemgrades[3][k]
                                        transcript_temp.iloc[j, 8] = studentexamgrades[3][k]

        column = i

 for i in range(9, 27):
     if isinstance(transcript_temp.iloc[i, 1], str) is True and isinstance(transcript_temp.iloc[i, 8], str) is True:
         print(transcript_temp.iloc[i, 0] + " was carried by student all the way through.")
     elif isinstance(transcript_temp.iloc[i, 1], str) is False and isinstance(transcript_temp.iloc[i, 8], str) is True:
         print("Student changed to " + transcript_temp.iloc[i, 0])
         for k in range(1, 8):
            while isinstance(transcript_temp.iloc[i, k], str) is False:
                transcript_temp.iloc[i, k] = "-"
     elif isinstance(transcript_temp.iloc[i, 1], str) is True and isinstance(transcript_temp.iloc[i, 8], str) is False:
         print("Student dropped " + transcript_temp.iloc[i, 0])
         newvar = 8
         while newvar > 0:
             transcript_temp.iloc[i, newvar] = ""
             newvar = newvar - 1
     else:
         pass



 transcript_temp.to_excel("output.xlsx")




