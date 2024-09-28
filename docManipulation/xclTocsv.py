import os
import pandas as pd

print('Enter path: \n')

dirPath = input()

cwd = os.getcwd()

files = [os.path.join(cwd, f) for f in os.listdir(cwd) if 
         os.path.isfile(os.path.join(cwd, f)) and f.endswith(".xlsx")]

newPath = os.path.join(cwd, r'newFiles')

if not os.path.exists(newPath):
    print("creating new path: ", newPath)
    os.makedirs(newPath)
    print("new path created.")
else:
    print("path already exists", newPath)

fileCount = 0

for file in files:
    head, tail = os.path.split(file)
    print("Reading file: ", tail)            
    read_file = pd.read_excel(file)

    newfilePath = newPath.join(newPath, tail)
    
    print("Converting file...")
    read_file.to_csv(newfilePath, index = None, header = True)
    print("File saved.")

    fileCount += 1

print(fileCount, "files converted.")

print("=============")
print("  Finished. ")
print("=============")

