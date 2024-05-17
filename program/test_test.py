from tkinter import Tk, filedialog

root = Tk() 
root.withdraw() 

root.attributes('-topmost', True) 

excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]) 
print(excel_file)
