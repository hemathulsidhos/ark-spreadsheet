import pandas as pd
import os
import shutil
import zipfile

#Batch download ARK metadata in CSV: arkimedes batch-download - EZID username EZID password --batch-format csv --batch-args "&column=dc.creator&column=dc.title&column=dc.date&column=dc.publisher&column=dc.type&column=_target"

#Note: Rename "filename" with the downloaded file name(5 instances).
#Note: Replace file location (C:/Users/hemat/Documents/Notepad++) based on your system settings.

#extract file


with zipfile.ZipFile("b9JRTe55Bv94as0f.zip","r") as zip_ref:
    zip_ref.extractall("C:/Users/hemat/Documents/Notepad++")


#CSV to Excel

read_file = pd.read_csv (r'C:\Users\hemat\Documents\Notepad++\b9JRTe55Bv94as0f.csv')
read_file.to_excel (r'C:\Users\hemat\Documents\Notepad++\b9JRTe55Bv94as0f.xlsx', index = None, header=True)

#delete CSV file

file = 'b9JRTe55Bv94as0f.csv'
if(os.path.exists(file) and os.path.isfile(file)):
  os.remove(file)
  print("CSV file deleted")
else:
  print("file not found")

#Rename file

source = 'C:/Users/hemat/Documents/Notepad++/b9JRTe55Bv94as0f.xlsx'
destination = 'C:/Users/hemat/Documents/Notepad++/ARK_Project.xlsx'

os.rename(source, destination)

path = 'ARK_Project.xlsx'
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

writer.save()
