import tkinter as tk
from openpyxl import Workbook

# Validation for Name
def is_valid_name(name):
    return name.isalpha()
# Validation for Roll Number
def is_valid_rollnumber(rollnumber):
    return rollnumber.isalnum()

def Attedence():
    name = name_entry.get()
    rollnumber=rollnumber_entry.get()
    if name and is_valid_rollnumber(rollnumber) and is_valid_name(name):
        worksheet.append([name,rollnumber])
        workbook.save("Attendence_DS.xlsx")
        name_entry.delete(0, "end")
        rollnumber_entry.delete(0, "end")
    else:
        validation_label.config(text="Invalid input. Check your Name or RollNumber.")
    
        
root = tk.Tk()
root.title("Add Attendence to Excel")

# Create an Excel workbook and add a worksheet
workbook = Workbook()
worksheet = workbook.active

# Create a label and an entry widget
name_label = tk.Label(root, text="Enter a Name:")
rollnumber_label=tk.Label(root, text="Enter a Roll Number")

name_label.grid(row=0, column=0, padx=10, pady=10)
rollnumber_label.grid(row=1, column=0, padx=10, pady=10)

name_entry = tk.Entry(root)
rollnumber_entry=tk.Entry(root)

name_entry.grid(row=0, column=1, padx=10, pady=10)
rollnumber_entry.grid(row=1, column=1, padx=10, pady=10)

# Create a button to add the name to Excel
add_button = tk.Button(root, text="Add Attendence", command=Attedence)
add_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Create a label for validation messages
validation_label = tk.Label(root, text="", fg="red")
validation_label.grid(row=3, columnspan=2, padx=10, pady=10)

root.mainloop()
