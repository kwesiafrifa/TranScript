from pathlib import Path

import tkinter as tk
from tkinter import filedialog
from tkinter import *
from traverse_folders import *

OptionList = ["2020",
              "2021"]


def browse_button():
    global filename
    filename = filedialog.askdirectory()


def create_transcripts():
    TestFunction(Path(filename).glob('**/*'))


root = Tk()
root.geometry('500x250')
root.title("TranscriptCreator 1.0")

variable = tk.StringVar(root)
variable.set(OptionList[0])

label = Label(root, text="Select folder containing student reports.")
btn_browse = Button(text=" Browse.. ", command=browse_button)
label1 = Label(root, text="Choose graduating class of students. ")
opt = tk.OptionMenu(root, variable, *OptionList)
opt.config(font=('Helvetica', 9))
label2 = Label(root, text="Enter number of students in desired graduating class.")
text = Text(root, height=1, width=4)
btn_create = Button(bg='blue', command=create_transcripts, text="Create transcripts", width=20, relief=RIDGE, fg='white')

label.place(relx=0, x=5, y=30)
label.config(font=("Helvetica", 12))
btn_browse.place(relx=0, x=350, y=30)
btn_browse.config(font=("Helvetica", 12))
label1.place(relx=0, x=5, y=70)
label1.config(font=("Helvetica", 12))
opt.place(relx=0, x=350, y=70)
opt.config(font=("Helvetica", 12))
label2.place(relx=0, x=5, y=110)
label2.config(font=("Helvetica", 10))
text.place(relx=0, x=350, y=110)
text.config(font=("Helvetica", 12))
btn_create.place(relx=0, x=150, y=200)
btn_create.config(font=("Helvetica", 12))

root.mainloop()

























































































































































































root = Tk()

