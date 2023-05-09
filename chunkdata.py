import pandas as pd

# Read the data from the Excel file
df = pd.read_excel('CLV_MTA01_Infinity_01_20230501_7136_Tubes.xlsx')

# Split the data into groups of 200 rows
groups = [df.iloc[i:i+200] for i in range(0, len(df), 200)]

# Save the first group to an Excel file with sheet name "E-Manifest"
filename = 'group_1.xlsx'
groups[0].to_excel(filename, index=False, sheet_name='E-Manifest')
print(f'Saved {len(groups[0])} rows to {filename}')

# Save the remaining groups to Excel files with the first three rows of the first group added to the beginning
for i, group in enumerate(groups[1:], start=2):
    filename = f'group_{i}.xlsx'
    group = pd.concat([groups[0].head(2), group])
    group.to_excel(filename, index=False, sheet_name='E-Manifest')
    print(f'Saved {len(group)} rows to {filename}')
