# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
import re
import msilib
os.chdir("C:\\New folder")
l=[]
for file in os.listdir():
    # print(os.getcwd())
    # # print(file)
    if (file.startswith("ZIC-")) or (file.startswith("ZCX-")) and (file.endswith(".mst")):
        l.append(file)
    elif (file.startswith("ZIC-")) or (file.startswith("ZCX-")) and (file.endswith(".msi")):
        l.append(file)
    elif (file.startswith("ZIC-")) or (file.startswith("ZCX-")) and (file.endswith(".exe")):
        l.append(file)
    else:
        pass
l=sorted(l)
for i in l:
    print(i)


        # MSI="msi or exe is NO FOUND..!"
# if (MSI.endswith(".msi")):










