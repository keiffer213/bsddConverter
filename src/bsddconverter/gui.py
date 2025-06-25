import os
import sys
from tkinter import Tk, Frame, Label, Entry, Button, Checkbutton, BooleanVar, END, filedialog, messagebox
from bsddconverter.mapper import run_excel2bsdd_conversion

HERE = os.path.dirname(__file__)
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, os.pardir, os.pardir))

def select_file(entry):
    """
    Opens a file dialog and returns the selected file path for the Excel file.
    """
    filename = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        entry.delete(0, END)
        entry.insert(0, filename)

def select_template_file(entry):
    """
    Opens a file dialog and returns the selected file path for the JSON template.
    """
    filename = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if filename:
        entry.delete(0, END)
        entry.insert(0, filename)

def select_output_file(entry):
    """
    Opens a file dialog and returns the selected file path for the output JSON file.
    """
    filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if filename:
        entry.delete(0, END)
        entry.insert(0, filename)

def run_converter(excel_entry, template_entry, output_path_entry, nulls_var):
    """
    Runs the conversion process when the Run Converter button is clicked.
    """
    excel_path = excel_entry.get().strip()
    template_path = template_entry.get().strip()
    output_path = output_path_entry.get().strip()
    without_nulls = nulls_var.get()

    if not excel_path:
        messagebox.showerror("Missing Info", "Please select an Excel file.")
        return
    if not template_path:
        messagebox.showerror("Missing Info", "Please select a JSON template.")
        return
    if not output_path:
        messagebox.showerror("Missing Info", "Please provide an output path.")
        return

    try:
        run_excel2bsdd_conversion(excel_path, template_path, output_path, remove_nulls=without_nulls)
        messagebox.showinfo("Success", f"JSON file saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def main():
    """Main function for the GUI."""
    app = Tk()
    app.title("Excel2bSDD Converter")
    app.geometry("600x225")

    mainframe = Frame(app)
    mainframe.pack(padx=20, pady=25)

    Label(mainframe, text="Excel File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    excel_entry = Entry(mainframe, width=60)
    excel_entry.grid(row=0, column=1, padx=5, pady=5)
    Button(mainframe, text="Browse", command=lambda: select_file(excel_entry)).grid(row=0, column=2, padx=5)

    Label(mainframe, text="JSON Template:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    template_entry = Entry(mainframe, width=60)
    template_entry.grid(row=1, column=1, padx=5, pady=5)
    Button(mainframe, text="Browse", command=lambda: select_template_file(template_entry)).grid(row=1, column=2, padx=5)

    Label(mainframe, text="Output Name:").grid(row=2, column=0, sticky="e", pady=5)
    output_name_entry = Entry(mainframe, width=60)
    output_name_entry.grid(row=2, column=1)
    Button(mainframe, text="Browse", command=lambda: select_output_file(output_name_entry)).grid(row=2, column=2, padx=5)

    nulls_var = BooleanVar()
    Checkbutton(mainframe, text="Remove nulls", variable=nulls_var).grid(row=3, column=1, sticky="w")

    Button(mainframe, text="Run Converter", bg="green", fg="white", 
           command=lambda: run_converter(excel_entry, template_entry, output_name_entry, nulls_var)
           ).grid(row=5, column=1, pady=20)

    app.mainloop()


if __name__ == "__main__":
    main()