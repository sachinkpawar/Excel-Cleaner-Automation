import pandas as pd
from tkinter import Tk, filedialog, messagebox

root = Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not file_path:
    messagebox.showinfo("Excel Cleaner", "No file selected.")
    exit()

df = pd.read_excel(file_path)

rows_before = len(df)

df = df.dropna(how="all")
df = df.drop_duplicates()

df.columns = df.columns.astype(str).str.strip()

for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].astype(str).str.strip()

rows_after = len(df)
rows_removed = rows_before - rows_after

output_file = file_path.replace(".xlsx", "_cleaned.xlsx")

df.to_excel(output_file, index=False)

messagebox.showinfo(
    "Cleaning Completed",
    f"Rows before: {rows_before}\n"
    f"Rows after: {rows_after}\n"
    f"Rows removed: {rows_removed}\n\n"
    f"Saved as:\n{output_file}"
)
