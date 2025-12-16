import pandas as pd

df = pd.read_excel("sample_data.xlsx", engine="openpyxl")

print("Rows BEFORE:", len(df))
print("Missing Sales BEFORE:", df["Sales"].isna().sum())

# Make Sales numeric (turns text/blank into NaN)
df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")

# Remove duplicates based on Email (more realistic than whole-row duplicates)
if "Email" in df.columns:
    df = df.drop_duplicates(subset=["Email"])
else:
    df = df.drop_duplicates()

# Fill missing sales with 0
df["Sales"] = df["Sales"].fillna(0)

print("Rows AFTER:", len(df))
print("Missing Sales AFTER:", df["Sales"].isna().sum())

# Summary
summary = pd.DataFrame({
    "Total Sales": [df["Sales"].sum()],
    "Number of Clients": [len(df)],
    "Average Sales": [df["Sales"].mean()],
})

with pd.ExcelWriter("output/summary.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Cleaned Data")
    summary.to_excel(writer, index=False, sheet_name="Summary")

print("Automation complete. Output saved to output/summary.xlsx")
