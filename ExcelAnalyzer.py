import pandas as pd
from tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu, messagebox

def load_excel_file():
    """Loads the Excel or CSV file and initializes the column dropdowns."""
    file_path = file_path_var.get()
    if not file_path:
        messagebox.showwarning("No File Path", "Please enter a valid file path.")
        return

    global df

    try:
        # Load the file into a DataFrame
        if file_path.lower().endswith(".csv"):
            df = pd.read_csv(file_path)
        elif file_path.lower().endswith((".xls", ".xlsx")):
            df = pd.read_excel(file_path)
        else:
            messagebox.showerror("Unsupported File", "The selected file type is not supported.")
            return

        if df.empty:
            messagebox.showerror("Empty File", "The file contains no data.")
            return

        # Populate column dropdowns with column names
        column1_dropdown['menu'].delete(0, 'end')
        column2_dropdown['menu'].delete(0, 'end')

        for col in df.columns:
            column1_dropdown['menu'].add_command(label=col, command=lambda value=col: column1_var.set(value))
            column2_dropdown['menu'].add_command(label=col, command=lambda value=col: column2_var.set(value))

        column1_var.set("Select Column")
        column2_var.set("Select Column")
        messagebox.showinfo("File Loaded", "File loaded successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

def update_values(column_var, value_dropdown, value_var):
    """Updates the values dropdown based on the selected column."""
    selected_column = column_var.get()
    if selected_column == "Select Column":
        value_dropdown['menu'].delete(0, 'end')
        value_var.set("Select Value")
        return

    try:
        unique_values = df[selected_column].dropna().unique()
        value_dropdown['menu'].delete(0, 'end')
        for val in unique_values:
            value_dropdown['menu'].add_command(label=val, command=lambda value=val: value_var.set(value))
        value_var.set("Select Value")
    except KeyError:
        messagebox.showerror("Error", "Failed to retrieve unique values for the selected column.")

def filter_data():
    """Filters the data based on the selected columns and values."""
    selected_column1 = column1_var.get()
    selected_value1 = value1_var.get()
    selected_column2 = column2_var.get()
    selected_value2 = value2_var.get()

    if any(selection == "Select Column" or value == "Select Value" for selection, value in 
           [(selected_column1, selected_value1), (selected_column2, selected_value2)]):
        messagebox.showwarning("Invalid Selection", "Please make valid selections for both filters.")
        return

    try:
        filtered_df = df[df[selected_column1] == selected_value1]
        count = filtered_df[filtered_df[selected_column2] == selected_value2].shape[0]
        messagebox.showinfo("Result", f"Number of records: {count}")
    except KeyError:
        messagebox.showerror("Error", "Invalid column or value selected.")

# Initialize GUI
root = Tk()
root.title("Excel Data Analyzer")

# Variables
df = pd.DataFrame()
file_path_var = StringVar()
column1_var = StringVar(value="Select Column")
value1_var = StringVar(value="Select Value")
column2_var = StringVar(value="Select Column")
value2_var = StringVar(value="Select Value")

# Widgets
Label(root, text="Enter File Path:").pack(pady=10)
Entry(root, textvariable=file_path_var, width=50).pack(pady=5)
Button(root, text="Load File", command=load_excel_file).pack(pady=5)

# Filter 1
Label(root, text="Select Filter 1 (e.g., Institution):").pack(pady=10)
column1_dropdown = OptionMenu(root, column1_var, "Select Column")
column1_dropdown.pack(pady=5)

value1_dropdown = OptionMenu(root, value1_var, "Select Value")
value1_dropdown.pack(pady=5)

# Filter 2
Label(root, text="Select Filter 2 (e.g., Role):").pack(pady=10)
column2_dropdown = OptionMenu(root, column2_var, "Select Column")
column2_dropdown.pack(pady=5)

value2_dropdown = OptionMenu(root, value2_var, "Select Value")
value2_dropdown.pack(pady=5)

# Update values when a column is selected
column1_var.trace("w", lambda *args: update_values(column1_var, value1_dropdown, value1_var))
column2_var.trace("w", lambda *args: update_values(column2_var, value2_dropdown, value2_var))

Button(root, text="Get Count", command=filter_data).pack(pady=20)

# Start the GUI loop
root.mainloop()
