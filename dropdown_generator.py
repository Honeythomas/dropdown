import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog

def load_excel_data():
    # Open a file dialog to select the Excel file
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    # Load the selected Excel file
    workbook = openpyxl.load_workbook(filepath)

    # Select the desired worksheet
    worksheet = workbook['Class 6']

    # Read the options from the specified columns in the worksheet
    curriculum_options = []
    class_options = []
    subject_options = []
    textbook_options = []
    strand_options = []

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        curriculum_options.append(row[0])
        class_options.append(row[1])
        subject_options.append(row[2])
        textbook_options.append(row[3])
        strand_options.append(row[4])

    # Update the dropdown options
    curriculum_dropdown['values'] = curriculum_options
    class_dropdown['values'] = class_options
    subject_dropdown['values'] = subject_options
    textbook_dropdown['values'] = textbook_options
    strand_dropdown['values'] = strand_options

# Create the UI
root = tk.Tk()
root.title("Excel Data to Dropdown UI")

# Create labels for dropdown boxes
curriculum_label = tk.Label(root, text="Curriculum:")
curriculum_label.pack()

class_label = tk.Label(root, text="Class:")
class_label.pack()

subject_label = tk.Label(root, text="Subject:")
subject_label.pack()

textbook_label = tk.Label(root, text="Textbook:")
textbook_label.pack()

strand_label = tk.Label(root, text="Strand:")
strand_label.pack()

# Create dropdown boxes
curriculum_dropdown = ttk.Combobox(root)
curriculum_dropdown.pack()

class_dropdown = ttk.Combobox(root)
class_dropdown.pack()

subject_dropdown = ttk.Combobox(root)
subject_dropdown.pack()

textbook_dropdown = ttk.Combobox(root)
textbook_dropdown.pack()

strand_dropdown = ttk.Combobox(root)
strand_dropdown.pack()

# Create a button to load the Excel data
load_button = tk.Button(root, text="Load Excel Data", command=load_excel_data)
load_button.pack()

# Run the UI
root.mainloop()