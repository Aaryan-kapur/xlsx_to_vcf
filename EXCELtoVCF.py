import os
import pip
package_names=['xlrd','openpyxl','pandas']
for i in package_names:
    try:
        pip.main(['install',i])
    except:
        print("Package "+i+" not found!")
        continue
import xlrd
import pandas as pd



def create_vcf(file,sheetName):
    print(file,sheetName)
    os.chdir("Excel")
    excelfile= pd.ExcelFile(file)
    column = excelfile.parse(sheetName)
    s = ""
    begin = "BEGIN:VCARD\nVERSION:2.1"
    for i in range(len(column)):
        fName=""
        sName=""
        secMail=""
        if(str(column["Phone"][i])!="nan"):
            if(str(column["First"][i])!="nan"):
                fName=str(column["First"][i])
            if(str(column["Last"][i])!="nan"):
                sName=str(column["Last"][i])        
            secN="\nN:"+ sName + ";" + fName + ";;;"
            secFN="\nFN:" + fName +" "+ sName
            secPhone="\nTEL;CELL:+"+str(column["Phone"][i]).split(".")[0]
            if("Mail" in column.columns.values):
                secMail=""
                if(str(column["Mail"][i]) != "nan"):
                    secMail="\nEMAIL;HOME:"+str(column["Mail"][i])
            s+=begin+secN + secFN +secPhone + secMail +"\nEND:VCARD\n"
    os.chdir("../Exported")
    text_file = open("Exported_"+file.split(".")[0] + "_"+sheetName+".vcf", "w")
    text_file.write(s)
    text_file.close()
    print("Completed!")
    os.chdir("..")




for file in os.listdir("Excel"):
    if(file.endswith(".xlsx")):
        print("Processing: "+file)
        for sheet in pd.ExcelFile("Excel/"+file).sheet_names:
            create_vcf(file,sheet)
        print("Completed!")
    else:
        print("File not supported: "+file)
        continue

# file= 'Contacts.xlsx'
# sheetName = "Contacts"
# create_vcf(file,sheetName)