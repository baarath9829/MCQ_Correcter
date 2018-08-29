import openpyxl
wb = openpyxl.load_workbook("MCQ_OMR.xlsx")
master_key = []
stringbuffer = ""
mark = 0
noofquestions = input("enter the number of questions :")
print (wb.sheetnames)
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
        
print (master_key)

for sheet in sheets:
    print (sheet)
    if sheet is not "answerkey":
        answersheet = wb[sheet]
        for i in range(1,noofquestions+1):
            stringbuffer = stringbuffer + str(answersheet["A" + str(i)].value)
            stringbuffer = stringbuffer + str(answersheet["B" + str(i)].value)
            stringbuffer = stringbuffer + str(answersheet["C" + str(i)].value)
            stringbuffer = stringbuffer + str(answersheet["D" + str(i)].value)
            print stringbuffer + "-->" + master_key[i-1]
            if stringbuffer is master_key[i-1]:
                mark = mark + 1
            stringbuffer = ""
        answersheet["F1"].value = "Total :"
        answersheet["G1"].value = mark
        wb.save("MCQ_OMR.xlsx")
        mark = 0
        
                
