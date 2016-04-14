#imports
from tkinter import *
from tkinter import ttk
import threading
import xlrd #module to read excel file
from decimal import * #decimal module for precise floating point calculation
import time #time module(used here for the 2 second wait)

running = True #Thread terminator
#data list and sum list
entries = [None] * 30 #list to hold the data read from excel in str data type 
sumEntries = [None]*8 #list to hold the summed values in float type
master = Tk()
master.title("Excel Reader")
master.configure(background="#565f6f")
master.wm_iconbitmap("favicon.ico")
master.resizable(0,0)
startButton = Button(master, text="START",bg="green",fg="white",width=20,font=("Helvetica",10))
startButton.grid(column=1, row=1)
btSep= Label(master,bg="#737373",width=40).grid(column=2,row=1)
stopButton = Button(master, text="STOP",bg="red",fg="black",width=20,font=("Helvetica",10))
stopButton.grid(column=3, row=1)
#Entry widgets to hold 30 values read from excel sheet
dispData = [None]*30
for i in range(30):
    dispData[i] = Entry(master,bd=1,relief=GROOVE,fg="#737373")
    dispData[i].grid(column=1,row=i+2,pady=2) 
    dispData[i].configure(highlightcolor="#000000",highlightthickness=0.2)
#end of data widgets
#mid column widget
sumLabel = Label(master,fg="#8fa9b9",bg="#565f6f",text="Maximum is",font=("Helvetica",10)).grid(column=2,row=16)
#end of mid column widget
#sum widget
dispSum = Entry(master,fg="#737373")
dispSum.grid(column=3,row=16)
dispSum.configure(highlightcolor="#000000",highlightthickness=2)
#end of sum widget


def readData():
    try:
        excelFile = xlrd.open_workbook('File.xls') #change the name & path of the excel file as per your requirement
        wsheet = excelFile.sheet_by_name('Sheet1') #change the sheet name if you need to.
        if wsheet.cell(8, 0).ctype != 2: #checks if the cell content is a number (value 2 means the content is a number. the values for text, date and other data can be obtained from xlrd documentation)
            return(0) #if not a number, return 0
        else:
            value = Decimal(wsheet.cell(8, 0).value) # if number, assign the value in the cell to variable named value
            return('%30f'%value)# return the value with 6 decimal digit precision
    except:
        return("N+") # if the cell is empty return a unique string N+
    
def processThread():
    i=0
    res_val = 0
    while sumEntries[7]== None: # outer loop
            while (entries[29] == None):# inner loop
                for j in range(30): # loop to assign 30 values read from excel to list
                    res_val = readData()
                    if((res_val !="N+")and running): #check if cell is not empty
                        entries[j] = readData() #assign values to list.
                        dispData[j].insert(1,(entries[j]))#print list values
                        time.sleep(2)#wait for 2 seconds
                    else:
                        break
                if(res_val =="N+"):#break loops if cell is empty
                    break
            
            if entries[29] != None: # the summing part of the loop
                floatEntries = map(float,entries) # convert string type list to float
                sumEntries[i] = sum(floatEntries) # call the inbuilt sum() if list[29] is not empty
                strSum = '%30f'%sumEntries[i]
                dispSum.delete(1,END)
                dispSum.insert(1,str(strSum))#print the sum list
                i=i+1 #increament index for sum list
                entries[29] = None #release the content from list[29] and make it none
            elif res_val == "N+": # break the outer loop if cell is empty
                break
    
        

t1_stop= threading.Event() #thread to process data
t1 = threading.Thread(target=processThread)

def start():
    global running
    running = True
    dispSum.delete(1, END)
    for j in range(30):
        dispData[j].delete(1,END)
    try:
        t1.start()
    except:
        pass
def stop():
    global running
    running = False
    
startButton.configure(command = start)
stopButton.config(command = stop)
    

master.mainloop()
