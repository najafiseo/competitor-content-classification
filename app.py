import pandas as pd
import os

directory_path = os.getcwd()
excel_files = [file for file in os.listdir(directory_path) if file.endswith('.xlsx')]
integrated_data = pd.DataFrame()
columns_to_select = ["Address", "Title ۱", "Meta Description ۱", "H1-۱", "H2-۱", "H2-۲", "Word Count", "Link Score", "Unique Inlinks"]

for file in excel_files:
    file_path = os.path.join(directory_path, file)
    data = pd.read_excel(file_path, usecols=columns_to_select) 
    data['Title ۱'].fillna('', inplace=True)
    integrated_data = integrated_data.append(data, ignore_index=True)


integrated_data.to_excel('integrated_data.xlsx', index=False)

xlsx_file_path = 'integrated_data.xlsx' 
df = pd.read_excel(xlsx_file_path)

num_keywords = int(input("Enter the number of keywords to filter: "))
filtered_dfs = {}
for i in range(num_keywords):
    keyword = input(f"Enter keyword {i + 1}: ")
    filtered_df = df[df['Title ۱'].str.contains(keyword, na=False)]  
    filtered_dfs[keyword] = filtered_df

excel_writer = pd.ExcelWriter('content classification.xlsx', engine='xlsxwriter')
for keyword, filtered_df in filtered_dfs.items():
    filtered_df.to_excel(excel_writer, sheet_name=keyword, index=False)

excel_writer.save()
os.remove('integrated_data.xlsx')

print("**********************Finished**********************")
