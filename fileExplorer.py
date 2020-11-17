import tkinter as tk
from tkinter import filedialog

temp = tk.Tk()
temp.withdraw()

file_path = filedialog.askopenfilename()

print(file_path)
input("press any key to exit")