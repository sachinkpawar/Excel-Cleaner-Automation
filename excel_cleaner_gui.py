import pandas as pd
from tkinter import Tk, filedialog

root = Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not file_path:
    print("No file selected.")
    exit()

df = pd.read_excel(file_path)

rows_before = len(df)

df = df.dropna(how="all")
df = df.drop_duplicates()

df.columns = df.columns.astype(str).str.strip()

for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].astype(str).str.strip()

rows_after = len(df)

output_file = file_path.replace(".xlsx", "_cleaned.xlsx")

df.to_excel(output_file, index=False)

print("Cleaning completed!")
print("Rows removed:", rows_before - rows_after)
print("Saved:", output_file)