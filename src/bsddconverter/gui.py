
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import os
from bsddconverter.mapper import run_excel2bsdd_conversion

HERE = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(HERE, os.pardir, 'templates', 'bsdd_import_model.json')

def select_file(entry):
    filename = filedialog.askopenfilename()
    if filename:
        entry.delete(0, END)
        entry.insert(0, filename)

def select_output_file(entry):
    filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if filename:
        entry.delete(0, END)
        entry.insert(0, filename)

def run_converter():
    excel_path = excel_entry.get()
    template_path = TEMPLATE_PATH
    output_path = f"./{output_name_entry.get()}.json"

    without_nulls = nulls_var.get()

    if not (excel_path and template_path and output_path):
        messagebox.showerror("Missing Info", "Please fill in all fields.")
        return

    try:
        run_excel2bsdd_conversion(excel_path, template_path, output_path, remove_nulls=without_nulls)
        messagebox.showinfo("Success", f"JSON file saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI setup
app = Tk()
app.title("Excel2bSDD Converter")
app.geometry("550x225")

mainframe = Frame(app)
mainframe.grid(column=0, row=0, sticky="nsew", padx=50, pady=50)

Label(app, text="Excel File:").grid(row=0, column=0, sticky="e")
excel_entry = Entry(app, width=60)
excel_entry.grid(row=0, column=1, padx=5, pady=5)
Button(app, text="Browse", command=lambda: select_file(excel_entry)).grid(row=0, column=2)

Label(app, text="Output Name:").grid(row=2, column=0, sticky="e")
output_name_entry = Entry(app, width=60)
output_name_entry.grid(row=2, column=1)

nulls_var = BooleanVar()
Checkbutton(app, text="Remove nulls", variable=nulls_var).grid(row=3, column=1, sticky="w")

Button(app, text="Run Converter", command=run_converter, bg="green", fg="white").grid(row=5, column=1, pady=20)

app.mainloop()
