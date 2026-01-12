import pandas as pd

df = pd.read_excel('02C WE.xlsx')

with open('analysis_output.txt', 'w', encoding='utf-8') as f:
    f.write('=== EXCEL FILE: 02C WE.xlsx ===\n\n')
    f.write(f'Total Rows: {len(df)}\n')
    f.write(f'Total Columns: {len(df.columns)}\n\n')
    
    f.write('=== COLUMNS ===\n')
    for i, col in enumerate(df.columns):
        f.write(f'  {i+1}. {col}\n')
    
    f.write('\n=== GRADE DISTRIBUTION ===\n')
    grade_dist = df['Grade'].value_counts()
    for grade, count in grade_dist.items():
        f.write(f'  {grade}: {count}\n')
    
    f.write('\n=== STATUS DISTRIBUTION ===\n')
    status_dist = df['Status'].value_counts()
    for status, count in status_dist.items():
        f.write(f'  {status}: {count}\n')
    
    f.write('\n=== MARKS STATISTICS ===\n')
    f.write(f'Subject Marks - Min: {df["Subject Marks"].min()}, Max: {df["Subject Marks"].max()}, Avg: {df["Subject Marks"].mean():.2f}\n')
    f.write(f'Assessment Marks - Min: {df["Assessment Marks"].min()}, Max: {df["Assessment Marks"].max()}, Avg: {df["Assessment Marks"].mean():.2f}\n')
    f.write(f'Final Marks - Min: {df["Final Marks"].min()}, Max: {df["Final Marks"].max()}, Avg: {df["Final Marks"].mean():.2f}\n')
    
    f.write('\n=== ALL STUDENT DATA ===\n')
    for idx, row in df.iterrows():
        f.write(f'\n--- Student {idx+1} ---\n')
        for col in df.columns:
            f.write(f'  {col}: {row[col]}\n')

print('Analysis saved to analysis_output.txt')
