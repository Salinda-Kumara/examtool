import pandas as pd

# Read Excel with no header first
df_raw = pd.read_excel('report.xls', header=None)

with open('excel_structure.txt', 'w', encoding='utf-8') as f:
    f.write(f"Shape: {df_raw.shape}\n")
    f.write("="*120 + "\n")
    f.write("First 25 rows:\n")
    f.write("="*120 + "\n")
    
    for i in range(min(25, len(df_raw))):
        vals = []
        for j in range(min(14, df_raw.shape[1])):
            val = df_raw.iloc[i, j]
            if pd.isna(val):
                vals.append("--")
            else:
                vals.append(str(val)[:18])
        f.write(f"R{i:02d}: " + " | ".join(f"{v:18s}" for v in vals) + "\n")
    
    # Test different header rows
    f.write("\n" + "="*120 + "\n")
    f.write("Testing different header positions:\n")
    f.write("="*120 + "\n")
    
    for skip_rows in [12, 13, 14]:
        try:
            test_df = pd.read_excel('report.xls', skiprows=skip_rows, nrows=5)
            f.write(f"\nSkipping {skip_rows} rows:\n")
            f.write(f"  Columns: {test_df.columns.tolist()}\n")
            f.write(f"  Shape: {test_df.shape}\n")
            if len(test_df) > 0:
                f.write(f"  First row: {test_df.iloc[0].tolist()}\n")
        except Exception as e:
            f.write(f"  Error with skip={skip_rows}: {str(e)}\n")

print("Analysis complete! Check excel_structure.txt")
