from tkinter import *

root = Tk()

root.geometry("400x500")
root.title('Kent CSV Converter')

def importExcel():
    myLabel2 = Label(root, text = "It Worked")
    myLabel2.grid(row=1, column=2)

def exportCSV():
    print("hello")

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