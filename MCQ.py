import openpyxl
from Tkinter import *
import string

def onsubmit(event):
    filename = txtbox_filename.get()
    #validatefilename(filename)
    wb = openpyxl.load_workbook(filename)
    master_key = []
    stringbuffer = ""
    mark = 0
    noofquestions = txtbox_noofquestions.get()
    noofquestions = int(noofquestions)
    #validatenumber(noofquestions)
    sheets =  wb.sheetnames
    if ("answerkey" in sheets):
        keysheet  = wb["answerkey"]
        for i in range(1,noofquestions+1):
            stringbuffer = stringbuffer + str(keysheet["A" + str(i)].value)
            stringbuffer = stringbuffer + str(keysheet["B" + str(i)].value)
            stringbuffer = stringbuffer + str(keysheet["C" + str(i)].value)
            stringbuffer = stringbuffer + str(keysheet["D" + str(i)].value)
            master_key.append(stringbuffer)
            stringbuffer = ""
        

    for sheet in sheets:
        print (sheet)
        if (sheet != "answerkey"):
            answersheet = wb[sheet]
            for i in range(1,noofquestions+1):
                stringbuffer = stringbuffer + str(answersheet["A" + str(i)].value)
                stringbuffer = stringbuffer + str(answersheet["B" + str(i)].value)
                stringbuffer = stringbuffer + str(answersheet["C" + str(i)].value)
                stringbuffer = stringbuffer + str(answersheet["D" + str(i)].value)
                if (stringbuffer == master_key[i-1]):
                    mark = mark + 1
                stringbuffer = ""
            answersheet["F1"].value = "Total :"
            answersheet["G1"].value = mark
            print (mark)
            wb.save("MCQ_OMR.xlsx")
            mark = 0
            
def validatefilename(filename):
    if(filename != ""):
        if (type(filename)== str):
            if (".xlsx" in filename):
                return 1
            else:
                print ("extension is invalid")
        else:
            print ("filename should be string")
    else:
        print ("text box is empty")
        
       
def validatenumber(number):
    checklist = list(string.ascii_lowercase)
    checklist+list(string.ascii_uppercase)
    checklist+list(string.punctuation)
    if(number != ""):
        flag = True
        for letter in checklist:
            if(letter in number):
                flag = False
        if (flag):
            flag = False
            for letter in range(10):
                if(str(letter) in number):
                    flag = True
            if (flag):
                return 1
            else:
                print ("there is no number")
        else:
            print ("there should not any special character or alphabet")
    else:
        print("text box is empty")
    return 0
    
window =  Tk()
Label(window , text="file name with extension").grid(row=0 , sticky=W)
txtbox_filename = Entry(window, width=20)
txtbox_filename.grid(row=0,column=1, sticky=E)
Label(window , text="number of question").grid(row=1 , sticky=W)
txtbox_noofquestions = Entry(window, width=20)
txtbox_noofquestions.grid(row=1,column=1, sticky=E)
button = Button(window , text="submit")
button.grid(row=2,column=1, sticky=E)
button.bind("<Button-1>" , onsubmit)
window.mainloop()
        
                
