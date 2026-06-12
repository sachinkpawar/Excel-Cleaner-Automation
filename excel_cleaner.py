import pandas as pd

print("=== Excel Cleaner ===")

input_file = "input.xlsx"
output_file = "cleaned_output.xlsx"

df = pd.read_excel(input_file)

rows_before = len(df)

df = df.dropna(how="all")


df = df.drop_duplicates()


df.columns = df.columns.astype(str).str.strip()


for col in df.select_dtypes(include=["object", "string"]).columns:
    df[col] = df[col].astype(str).str.strip()

rows_after = len(df)
removed_rows = rows_before - rows_after

df.to_excel(output_file, index=False)

print(f"Rows before cleaning : {rows_before}")
print(f"Rows after cleaning  : {rows_after}")
print(f"Rows removed         : {removed_rows}")
print(f"Saved as             : {output_file}")