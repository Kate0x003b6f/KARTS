import csv
import glob
import os.path
import shutil
from datetime import *

line_item = []
processing = []
pline = []
counter1 = 0
shippingCharges = {}
discounts = {}
# Converts Squarespace to these headers format
headings = "TRANSACTION ID", "ACCOUNTING ID", "PURCHASE ORDER NUMBER", "PO DATE", "SHIP TO NAME", "SHIP TO ADDRESS 1", \
           "SHIP TO ADDRESS 2", "SHIP TO CITY", "SHIP TO STATE", "SHIP TO ZIP", "SHIP TO COUNTRY", "STORE NUMBER", \
           "BILL TO NAME", "BILL TO ADDRESS 1", "BILL TO ADDRESS 2", "BILL TO CITY", "BILL TO STATE", "BILL TO ZIP", \
           "BILL TO COUNTRY", "BILL TO CODE", "SHIP VIA", "SHIP DATE", "TERMS", "NOTE", "DEPARTMENT NUMBER", \
           "CANCEL DATE", "DO NOT SHIP BEFORE", "DO NOT SHIP AFTER", "Allowance Percent1", "Allowance Amount1", \
           "Allowance Precent2", "Allowance Amount2", "LINE #", "VENDOR PART #", "BUYERS PART #", "UPC #", "DESCRIPTION", \
           "QUANTITY", "UOM", "UNIT PRICE", "ITEM NOTES", "CUSTOMER ORDER #", "BUYER NAME", "BUYER PHONE", "BUYER FAX", \
           "BUYER EMAIL", "SHIP TO NAME", "INFO CONTACT NAME", "INFO CONTACT PHONE", "INFO CONTACT EMAIL", \
           "DEL CONTACT NAME", "DEL CONTACT PHONE", "DEL CONTACT EMAIL", "DEL REFERENCE", "SHIP TO ADDRESS 3", \
           "Item Notes", "SHIP TO ADDRESS 4", "ISA Timestamp", "Attention"

# creates a blank processing line
while counter1 < 59:
    pline.append("")
    counter1 += 1
lineNum = 1
# Makes a list of all the files in the folder
# folder_path = '\\\\tnt096\\Users\\tnt096\\Downloads'
folder_path = r'C:\\True Commerce\\Transaction Manager\\Export\\Web Orders'
file_type = '\*csv'
files = glob.glob(folder_path + file_type)
# Checks to make sure the file list is not empty
try:
    firstFile = files[0]
except IndexError:
    print("Unable to find file. Shutting down...")
    raise SystemExit
# Chooses the newest file in the folder
maxFile = max(files, key=os.path.getctime)
fileCreationDate = datetime.fromtimestamp(os.path.getctime(maxFile))
outputFile = "C:\\True Commerce\\Transaction Manager\\Export\\Web Order_" + fileCreationDate.strftime("%d_%H_%M") + ".csv"
output = open(outputFile, 'w', newline='')
writer = csv.writer(output)
fileName = str(os.path.basename(maxFile))
# Quits if newest file is not named orders
if fileName.find("orders") == -1:
    print("Unexpected file found. Please resolve.")
    raise SystemExit
# reads in the bookmark
bookmarkFile = open(r'C:\Scripts\Stored Data\bookmark.txt', "r")
bookmark = int(bookmarkFile.read())
bookmarkFile.close()

"""Main functionality begins here ----------------------------------------------------------------------------------"""
with open(maxFile, 'r', errors='replace') as PO:
    rcsv = csv.reader(PO, delimiter=',')
    counter2 = 0
    for line in rcsv:
        # skips any blank lines (sometimes squarespace export has blank lines at end of file)
        if line[0] == "":
            continue
        # if it's reading a line, and not the header, do the following:
        if counter2 > 0:
            if line[0] == pline[0]:
                lineNum += 1
                i = 0
                while i < len(line):
                    if line[i] == "":
                        line[i] == processing[-1][i]
            else:
                counter1 = 0
                while counter1 < 59:
                    pline[counter1] = ""
                    counter1 += 1
                lineNum = 1
            pline[0] = 850
            if line[sIndex["Checkout Form: Company Name"]].lower() == "personal":
                line[sIndex["Checkout Form: Company Name"]] = line[sIndex["Billing Name"]]
                pline[4] = line[sIndex["Billing Name"]]
            else:
                pline[4] = line[sIndex["Checkout Form: Company Name"]]
                pline[54] = line[sIndex["Shipping Name"]]
            companyCode = line[sIndex["Checkout Form: Company Name"]].replace("/", "").replace("\\", "")
            companyCode = companyCode.replace("#", "").replace("@", "").replace("*", "")
            pline[1] = companyCode[0:20]
            pline[2] = line[sIndex["Order ID"]]
            try:
                pline[3] = datetime.strptime(line[sIndex["Paid at"]], '%Y-%m-%d %H:%M:%S %z').strftime("%m/%d/%Y")
                pline[57] = datetime.strptime(line[sIndex["Paid at"]], '%Y-%m-%d %H:%M:%S %z').strftime("%m/%d/%Y %H:%M:%S")
            except:
                if len(line[sIndex["Paid at"]]) > 0:
                    if not line[sIndex["Paid at"]][0] == 0:
                        line[sIndex["Paid at"]] = "0" + line[3]
                    if line[sIndex["Paid at"]][4] == "/":
                        line[sIndex["Paid at"]] = line[sIndex["Paid at"]][0:3]+"0"+line[sIndex["Paid at"]][3:]
                    if len(line[sIndex["Paid at"]]) < 16:
                        line[sIndex["Paid at"]] = line[sIndex["Paid at"]][0:11] + "0" + line[sIndex["Paid at"]][11:]
                    line[sIndex["Paid at"]] = line[sIndex["Paid at"]] + ":00"
                    pline[3] = datetime.strptime(line[sIndex["Paid at"]], '%m/%d/%Y %H:%M:%S').strftime("%m/%d/%Y")
                    pline[57] = datetime.strptime(line[sIndex["Paid at"]], '%m/%d/%Y %H:%M:%S').strftime("%m/%d/%Y %H:%M:%S")
            pline[5] = line[sIndex["Shipping Address1"]]
            pline[6] = line[sIndex["Shipping Address2"]]
            pline[7] = line[sIndex["Shipping City"]]
            pline[8] = line[sIndex["Shipping Province"]]
            pline[9] = line[sIndex["Shipping Zip"]]
            pline[10] = "US"
            pline[11] = ""
            pline[12] = line[sIndex["Billing Name"]]
            pline[13] = line[sIndex["Billing Address1"]]
            pline[14] = line[sIndex["Billing Address2"]]
            pline[15] = line[sIndex["Billing City"]]
            pline[16] = line[sIndex["Billing Province"]]
            pline[17] = line[sIndex["Billing Zip"]]
            pline[18] = "US"
            pline[19] = ""
            pline[20] = "TRIPPNT GROUND"
            pline[21] = ""
            pline[22] = "CREDIT CARD"
            pline[23] = ""
            pline[24] = ""
            pline[25] = ""
            pline[26] = ""
            pline[27] = ""
            pline[28] = ""
            pline[29] = ""
            pline[30] = ""
            pline[31] = ""
            pline[32] = lineNum
            pline[33] = line[sIndex["Lineitem sku"]]
            pline[34] = line[sIndex["Lineitem sku"]]
            pline[35] = ""
            pline[36] = line[sIndex["Lineitem name"]]
            pline[37] = line[sIndex["Lineitem quantity"]]
            pline[38] = "Each"
            pline[39] = line[sIndex["Lineitem price"]]
            pline[40] = line[sIndex["Checkout Form: Note / Additional Info"]]
            pline[41] = line[sIndex["Checkout Form: Your Internal Purchase Order (if applicable)"]]
            pline[42] = line[sIndex["Billing Name"]]
            pline[43] = line[sIndex["Billing Phone"]]
            pline[44] = ""
            pline[45] = line[sIndex["Email"]]
            pline[46] = line[sIndex["Shipping Name"]]
            pline[47] = line[sIndex["Billing Name"]]
            pline[48] = line[sIndex["Billing Phone"]]
            pline[49] = line[sIndex["Email"]]
            pline[50] = line[sIndex["Shipping Name"]]
            pline[51] = line[sIndex["Shipping Phone"]]
            pline[52] = ""
            pline[53] = ""
            pline[54] += "/" + line[sIndex["Checkout Form: Your Internal Purchase Order (if applicable)"]]
            pline[55] = line[sIndex["Checkout Form: Shipping Instructions"]][0:50]
            if len(processing) > 0:
                if pline[2] == processing[-1][2]:
                    i = 0
                    while i < len(pline):
                        if pline[i] == "":
                            pline[i] = processing[-1][i]
                        i += 1
            processing.append(pline[0:59])
            # generates additional lines for shipping and discounts
            if not line[sIndex["Shipping"]] == "" or line[sIndex["Shipping"]] == "0.00":
                shipLine = pline[0:]
                shipLine[33] = "SHIPPING"
                shipLine[34] = "SHIPPING"
                shipLine[36] = ""
                shipLine[37] = 1
                shipLine[39] = line[sIndex["Shipping"]]
                shippingCharges[line[0]] = shipLine
            try:
                line[sIndex["Discount Amount"]] = float(line[sIndex["Discount Amount"]])
                discounts[line[0]] = line[sIndex["Discount Amount"]]
                discountLine = pline[0:59]
                discountLine[33] = "DISCOUNT"
                discountLine[34] = "DISCOUNT"
                discountLine[36] = ""
                discountLine[37] = 1
                discountLine[39] = -1 * line[sIndex["Discount Amount"]]
                discounts[line[0]] = discountLine
            except ValueError:
                continue
        else:
            # records the indexes of header fields (I had trouble with Squarespace adding columns at random)
            columnCounter = 0
            sIndex = {}
            for column in line:
                sIndex[column] = columnCounter
                columnCounter += 1
        counter2 += 1

writer.writerow(headings)
for item in shippingCharges.values():
    processing.append(item)
for item in discounts.values():
    if item[39] < 0:
        processing.append(item)
i = 0
lineItems = {}
for item in processing:
    if not item[2] in lineItems.keys():
        lineItems[item[2]] = 1
    else:
        lineItems[item[2]] += 1
    item[32] = lineItems[item[2]]
# Writes any orders  with order numbers greater than the bookmark
orderNums = []
for item in processing:
    orderNum = int(item[2].replace("-NP", ""))
    orderNums.append(orderNum)
    if orderNum > bookmark:
        writer.writerow(item)
output.close()
# Deletes blank input files and input files with no new orders
if len(orderNums) <= 0:
    os.remove(outputFile)
    print("File was blank")
newBookmark = max(orderNums)
# deletes output file if no new order found
if newBookmark <= bookmark:
    os.remove(outputFile)
# updates bookmark to newest order number
else:
    bookmarkFile = open(r'C:\Scripts\Stored Data\bookmark.txt', "w")
    bookmarkFile.write(str(newBookmark))
    bookmarkFile.close()
# renames the file and puts it in the processed folder
shutil.move(maxFile, "C:\\True Commerce\\Transaction Manager\\Export\\Web Orders\\Processed\\SQRSPCE"
            +fileCreationDate.strftime("_%d_%H_%M")+".csv")
