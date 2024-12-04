# Excel Data Analyzer

**Excel Data Analyzer** is a Python script that allows users to load Excel or CSV files, analyze data, and filter it dynamically through a graphical user interface (GUI). Built with `Tkinter` and `pandas`, this tool is ideal for users looking to quickly query records based on multiple criteria without needing advanced data analysis skills.

---

## Features

- **File Support**: Works with `.csv`, `.xls`, and `.xlsx` file formats.
- **Dynamic Dropdowns**: Automatically populates dropdown menus with column names and unique values from the dataset.
- **Hierarchical Filtering**:
  - Filter by **Institution Name** or similar column.
  - Further refine by selecting a **Role** or equivalent column.
- **Record Count**: Displays the number of records that match the selected criteria.
- **Error Handling**:
  - Handles invalid file types and empty datasets gracefully.
  - Provides user-friendly warnings for invalid selections.

---

## Requirements

Ensure you have the following installed:

- Python 3.7 or higher
- Required libraries:
  - `pandas`
  - `tkinter` (usually included in Python standard library)

Install required Python libraries using pip:

```bash
pip install pandas
