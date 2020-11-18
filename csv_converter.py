import tkinter as tk
from tkinter import filedialog
from tkinter import *
from numpy.core.records import array
import pandas as pd
import csv



root = tk.Tk()


root.geometry("400x500")
root.title('Kent CSV Converter')

imported = False

second_table = [["SAP Code","Qty","Description","Unit Price","Discount","Total Price","Material Grade","Finish","Est Eng Hours","Est Production Hours","Product Group 1",
                "Product Group 2", "Product Group 3", "Will need Sub Con", "Promise Date", "PDM Project", "Estimated Cost Price", "sap code"]]


def importExcel():
    file_path = filedialog.askopenfilename()
    data = pd.read_excel (file_path, sheet_name="Quote" , index_col=23,header=23)
# SECOND TABLE COLUMNS
    endOfSecondTable = data.iloc[0,8]

    for x in range(1000):
        if endOfSecondTable != "Total:":
            endOfSecondTable = data.iloc[x+1,8]

            second_table.append( [ data.iloc[x,0],data.iloc[x,1],data.iloc[x,2],data.iloc[x,7],data.iloc[x,8],data.iloc[x,9],data.iloc[x,11],data.iloc[x,12],data.iloc[x,13],
                             data.iloc[x,14],data.iloc[x,15],data.iloc[x,16],data.iloc[x,17],data.iloc[x,18],data.iloc[x,20],data.iloc[x,22],data.iloc[x,23] ] )
        else:
            break


    global imported
    imported = True
    importNotification = Label(root, text = "Your File has Been Imported!")
    importNotification.grid(row=1, column=2)
    importNotification.place(relx=0.5, rely=0.4, anchor=CENTER)
    exportNotification = Label(root, text = "                                                                                       ")
    exportNotification.grid(row=1, column=2)
    exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)

def exportCSV():
    if imported == True:
        second_table_dataframe = pd.DataFrame(second_table)
        second_table_dataframe.to_csv('quotation.csv')
        exportNotification = Label(root, text = "Your File has Been Exported To Your Desktop!")
        exportNotification.grid(row=1, column=2)
        exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)
    else:        
        exportNotification = Label(root, text = "You Have Not Imported A File Yet!")
        exportNotification.grid(row=1, column=2)
        exportNotification.place(relx=0.5, rely=0.6, anchor=CENTER)


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
exitAppButton.place(relx=0.5, rely=0.7, anchor=CENTER)



root.mainloop()