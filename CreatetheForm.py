import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook
import csv

# Load the Excel workbook
workbook = load_workbook('DropdownData.xlsx')
# Function to get values om a sheet
def get_values(sheet_name):
    sheet = workbook[sheet_name]
    return [cell.value for cell in sheet['A'][1:]]

# Get values for dropdowns
job_titles = get_values('Job Title')
departments = get_values('Department')
grades = get_values('Grade')

# Create the main window
root = tk.Tk()
root.title("Form")

# Load the company logo
logo = tk.PhotoImage(file="company_logo.png")  # Replace with your logo file path

# Create and place the logo
logo_label = ttk.Label(root, image=logo)
logo_label.grid(column=0, row=0, columnspan=2, padx=10, pady=10)

# Create and place the dropdowns
ttk.Label(root, text="Job Title:").grid(column=0, row=1, padx=10, pady=5)
job_title_var = tk.StringVar()
job_title_dropdown = ttk.Combobox(root, textvariable=job_title_var, values=job_titles)
job_title_dropdown.grid(column=1, row=1, padx=10, pady=5)

ttk.Label(root, text="Department:").grid(column=0, row=2, padx=10, pady=5)
department_var = tk.StringVar()
department_dropdown = ttk.Combobox(root, textvariable=department_var, values=departments)
department_dropdown.grid(column=1, row=2, padx=10, pady=5)

ttk.Label(root, text="Grade:").grid(column=0, row=3, padx=10, pady=5)
grade_var = tk.StringVar()
grade_dropdown = ttk.Combobox(root, textvariable=grade_var, values=grades)
grade_dropdown.grid(column=1, row=3, padx=10, pady=5)

# Create and place the checkbox
internal_var = tk.BooleanVar()
internal_checkbox = ttk.Checkbutton(root, text="Internal", variable=internal_var)
internal_checkbox.grid(column=0, row=4, columnspan=2, padx=10, pady=5)

# Function to save form data to a CSV file
def save_data():
    with open('form_data.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([job_title_var.get(), department_var.get(), grade_var.get(), internal_var.get()])

# Create and place the submit button
submit_button = ttk.Button(root, text="Submit", command=save_data)
submit_button.grid(column=0, row=5, columnspan=2, padx=10, pady=10)

# Run the application
root.mainloop()