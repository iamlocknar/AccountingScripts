#pip install xlrd

import xlrd
book=xlrd.open_workbook("TestData.xls")
book2=xlrd.open_workbook("170314.xlsx")
sheet=book.sheet_by_index(0)
sheet2=book2.sheet_by_index(0)
#Salesforce Deposit Sheet
DSSF = dict()
DSSFsum = float(0)
#Deposit Sheet Manual Entry
DSME = []
#Matching/Catching Dicts
notfoundDSSF = []
notfoundDSME = {}
#As the DSME does not have unique keys, possible duplicates may occur
#This dict will use the value amount as a key and a incrimented amount to show tally
possibleduplicates = {}
#Pull Data in dict and list
for row in range(sheet.nrows):
    #Skips Adding Header
    if row == 0:
        continue
    DSSFsum+=sheet.cell_value(row,0)
    DSSF.setdefault(sheet.cell_value(row,1), [])
    DSSF[sheet.cell_value(row,1)].append(float(sheet.cell_value(row,0)))
for row in range(sheet2.nrows):
    #skips adding Header
    if row == 0:
        continue
    if sheet2.cell_value(row,0) == '':
        break
    DSME.append(float(sheet2.cell_value(row,0)))

#Sort
DSME.sort()
DSSFCombined = {}
#Manipulated Dict duplicate
DSSFPoped = DSSF

#Chekc if Manual Entry is in Deposit Sheet, then pops a single entry from key list
for entry in DSME:
    for key in DSSF:
        if entry in DSSFPoped[key]:
            #pop uses index, remove uses first matching value
            DSSFPoped[key].remove(entry)
            break
        else:
            continue
        break



for key in DSSFPoped:
    if key == '':
        for item in DSSFPoped[key]:
            DSSFCombined.setdefault('No Check #', [])
            DSSFCombined['No Check #'].append(item)
    else:
        keysum=(round(sum(DSSFPoped[key]),2))
        if keysum > 0:
            DSSFCombined.setdefault(key, [])
            DSSFCombined[key].append(round(sum(DSSF[key]),2))


#if item is missing in DSME
for key in DSSFCombined:
    for item in DSSFCombined[key]:
        if item not in DSME:
            notfoundDSME.setdefault(key, [])
            notfoundDSME[key].append(item)



#TODO add the catch to find missing entries from Salesforce Deposit Sheet
#for entry in DSME:


#If no missing entries
if not notfoundDSME and not notfoundDSSF:
    print("Deposit Sheet and Manual Entry match!")
    print("Manual Deposit Sheet = ", sum(DSME))
    print("Deposit Sheet SF = ", (round(DSSFsum, 2)))

#if there are missing entries
if notfoundDSME or notfoundDSSF:
    print("These Items Are Not Found in Deposit Sheet Manual Entry:")
    for key in notfoundDSME:
        print("Check #: ", key,"    Amounts: ", DSSF[key])

    print("These Items Are Not Found in Salesforce Deposit:")
    for entry in notfoundDSSF:
        print("Amount: ", entry)
