import pandas as pd
import os
import shutil
import zipfile
import warnings
warnings.simplefilter(action='ignore',category=FutureWarning)
import subprocess

#Batch download ARK metadata in CSV: arkimedes batch-download - EZIDUsername EZIDPassword --batch-format csv --batch-args "&column=dc.creator&column=dc.title&column=dc.date&column=dc.publisher&column=dc.type&column=_target"

#arkimedes batch-download - EZIDUsername EZIDPassword --batch-format csv --batch-args "&column=dc.creator&column=dc.title&column=dc.date&column=dc.publisher&column=dc.type&column=_target" && python ark-spreadsheet-2.0.py

#***Make sure there is no zip and/or csv file in the directory_path***

#Note: Replace file location (directory_path) based on your system settings.

directory_path = r'C:\Users\hemat\Documents\VENV\pyvenv\Scripts'

# Navigate to the directory
subprocess.run(['cd', directory_path], shell=True, check=True)

# Rename the downloaded zip file to 'EZID.zip'
subprocess.run(['ren', '*.zip', 'EZID.zip'], shell=True, check=True)

# Extract the zip file
with zipfile.ZipFile("EZID.zip","r") as zip_ref:
    zip_ref.extractall(directory_path)

# Rename the extracted csv file to 'EZID.csv'
subprocess.run(['ren', '*.csv', 'EZID.csv'], shell=True, check=True)

#Convert CSV to Excel

csv_file_path = os.path.join(directory_path, 'EZID.csv')
read_file = pd.read_csv(csv_file_path)

excel_file_path = os.path.join(directory_path, 'EZID.xlsx')
read_file.to_excel(excel_file_path, index=None, header=True)

#delete CSV file

file = 'EZID.csv'
if(os.path.exists(file) and os.path.isfile(file)):
  os.remove(file)
  print("CSV file deleted")
else:
  print("file not found")

#Generate output file (ARK_Project_Output) with separate tabs

path = 'EZID.xlsx'
df = pd.read_excel(path)
df.head()

df1 = df[df['_target'].str.contains('aviary')]
#print(df1)

df2 = df[df['_target'].str.contains('digital|public')]
#print(df2)

df3 = df[df['_target'].str.contains('cardinal')]
#print(df3)

df4 = df[df['_target'].str.contains('ezid')]
#print(df4)

df5 = df[~df['_target'].str.contains('aviary|digital|public|cardinal|ezid')]
#print(df5)


writer = pd.ExcelWriter("ARK_Project_Output.xlsx", engine = 'xlsxwriter')

df1.to_excel(writer, sheet_name='aviary', index = False)
df2.to_excel(writer, sheet_name='islandora', index = False)
df3.to_excel(writer, sheet_name='finding aid', index = False)
df4.to_excel(writer, sheet_name='reuse', index = False)
df5.to_excel(writer, sheet_name='others', index = False)

writer.close()

#Delete EZID.xlsx 

file = 'EZID.xlsx'
if(os.path.exists(file) and os.path.isfile(file)):
  os.remove(file)
  print("XLSX file deleted")
else:
  print("file not found")


#Verify output file
file_name = "Ark_Project_Output.xlsx"
file_path = os.path.join(directory_path, file_name)

if os.path.exists(file_path):
    print(f"{file_name} file is available in the directory.")
else:
    print(f"{file_name} file is not found in the directory.")
