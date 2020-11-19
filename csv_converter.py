import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter.ttk import Progressbar
import time
from numpy.core.numeric import NaN, full
from numpy.core.records import array
import pandas as pd
import os



root = tk.Tk()


root.geometry("400x500")
root.title('Kent CSV Converter')

imported = False

try:
    os.remove("quotation.csv")
except:
    print("No Client Files to Delete")

try:
    os.remove("first_quotation.csv")
    os.remove("second_quotation.csv")
except:
    print("No Enterpise Files To Delete")



first_table = [["Quotation No","Revision","No","Currency","Contact","Company","Phone","Email","Products","Project","Delivery","Date","Originator","Manufacturing Time","Quote Valid for",
                "Market Sector","Civil or Building", "Site Fixings", "Engineer Head", "Business Unit", "Customer PO Number", "Delivery Address", "Delivery Address Complete Y/N"]]

second_table = [["SAP Code","Qty","Description","Unit Price","Discount","Total Price","Material Grade","Finish","Est Eng Hours","Est Production Hours","Product Group 1",
                "Product Group 2", "Product Group 3", "Will need Sub Con", "Promise Date", "PDM Project", "Estimated Cost Price", "sap code"]]


def importExcel():
    file_path = filedialog.askopenfilename()
    progress = Progressbar(root, orient = HORIZONTAL, 
        length = 100, mode = 'determinate') 
    progress.place(relx=0.5, rely=0.4, anchor=CENTER)
    progress['value'] = 20
    root.update_idletasks() 

    first_table_data = pd.read_excel (file_path, sheet_name="Quote" , index_col=None,header=None)
    second_table_data = pd.read_excel (file_path, sheet_name="Quote" , index_col=23,header=23)
    progress['value'] = 40
    root.update_idletasks() 

# FIRST TABLE COLUMNS
    first_table.append([ first_table_data.iloc[3,4], first_table_data.iloc[8,2],first_table_data.iloc[9,2],first_table_data.iloc[10,2],first_table_data.iloc[12,2],
                         first_table_data.iloc[13,2],first_table_data.iloc[14,2],first_table_data.iloc[15,2],first_table_data.iloc[16,2],first_table_data.iloc[12,7],first_table_data.iloc[13,7],
                         first_table_data.iloc[20,1],first_table_data.iloc[20,3],first_table_data.iloc[20,5],first_table_data.iloc[20,8],first_table_data.iloc[3,14],first_table_data.iloc[4,14],
                         first_table_data.iloc[5,14],first_table_data.iloc[6,14],first_table_data.iloc[7,14],first_table_data.iloc[8,14],first_table_data.iloc[5,20],
                         first_table_data.iloc[9,20],first_table_data.iloc[12,20] ])
# SECOND TABLE COLUMNS
    endOfSecondTable = second_table_data.iloc[0,8]

    for x in range(1000):
        if endOfSecondTable != "Total:":
            endOfSecondTable = second_table_data.iloc[x+1,8]

            second_table.append( [ second_table_data.iloc[x,0],second_table_data.iloc[x,1],second_table_data.iloc[x,2],second_table_data.iloc[x,7],second_table_data.iloc[x,8],
                             second_table_data.iloc[x,9],second_table_data.iloc[x,11],second_table_data.iloc[x,12],second_table_data.iloc[x,13],
                             second_table_data.iloc[x,14],second_table_data.iloc[x,15],second_table_data.iloc[x,16],second_table_data.iloc[x,17],
                             second_table_data.iloc[x,18],second_table_data.iloc[x,20],second_table_data.iloc[x,22],second_table_data.iloc[x,23] ] )
        else:
            break
    progress['value'] = 80
    root.update_idletasks() 

# PLACING NOTIFCATIONS
    global imported
    imported = True
    progress['value'] = 100
    progress.pack_forget
    importNotification = Label(root, text = "Your File has Been Imported!")
    importNotification.grid(row=1, column=2)
    importNotification.place(relx=0.5, rely=0.4, anchor=CENTER)
    exportNotification = Label(root, text = "                                                                                       ")
    exportNotification.grid(row=1, column=2)
    exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)

def exportCSV():
    if imported == True:
        exportNotification = Label(root, text = "Please Choose A Platform to Export to!")
        exportNotification.grid(row=1, column=2)
        exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)
        clientButton = Button(root, text = "For Client", font=('Helvetica 18 bold',10), command=clientExport, bg="blue",fg="white",padx=5,pady=5)
        clientButton.grid(row=1, column=2)
        clientButton.place(relx=0.3, rely=0.7, anchor=CENTER)
        enterpriseButton = Button(root, text = "For Enterprise", font=('Helvetica 18 bold',10), command=enterpriseExport, bg="blue",fg="white",padx=5,pady=5)
        enterpriseButton.grid(row=1, column=2)
        enterpriseButton.place(relx=0.7, rely=0.7, anchor=CENTER)
    else:        
        exportNotification = Label(root, text = "You Have Not Imported A File Yet!")
        exportNotification.grid(row=1, column=2)
        exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)

def clientExport():
    first_table_dataframe = pd.DataFrame(first_table)
    first_table_dataframe.to_csv('first_quotation.csv')
    second_table_dataframe = pd.DataFrame(second_table)
    second_table_dataframe.to_csv('second_quotation.csv')

def enterpriseExport():
    full_table = []
    full_table.append(first_table[0]+second_table[0])
    full_table.append(first_table[1]+second_table[1])
    for x in range(2,200):
        try:
            full_table.append(first_table[x]+second_table[x])
        except:
            print("Ran out of rows error")
    full_table_dataframe = pd.DataFrame(full_table)
    full_table_dataframe.to_csv('quotation.csv')



def exitButton():
    root.destroy()


appHeading = Label(root, text = "Kent Excel to CSV Converter", font=('Helvetica 18 bold',15))
importButton = Button(root, text = "Import File", font=('Helvetica 18 bold',25), command=importExcel,bg="green",fg="white",padx=20,pady=5)
exportButton = Button(root, text = "Export CSV", font=('Helvetica 18 bold',23), command=exportCSV, bg="blue",fg="white",padx=20,pady=5)
exitAppButton = Button(root, text = "Exit Application", font=('Helvetica 18 bold',19), command=exitButton, bg="red",fg="white",padx=20,pady=5)

appHeading.grid(row =0, column=0)
importButton.grid(row=1, column=0)
exportButton.grid(row=2, column=0)
exitAppButton.grid(row=3, column=0)

appHeading.place(relx=0.5, rely=0.1, anchor=CENTER)
importButton.place(relx=0.5, rely=0.3, anchor=CENTER)
exportButton.place(relx=0.5, rely=0.5, anchor=CENTER)
exitAppButton.place(relx=0.5, rely=0.9, anchor=CENTER)



root.mainloop()