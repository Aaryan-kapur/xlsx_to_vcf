import os

for i in os.listdir("Excel"):
    os.chdir("Excel")
    if(i.endswith(".xlsx")):
        print(i)
        j = i.replace(" ","").replace("_","").lower()
        os.rename(i,j)
        print("Renamed!")
    else:
        print("File not supported: "+i)
        continue
    os.chdir("..")
print("Completed!")