import csv
import glob
import os.path
import shutil
from datetime import *
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import ssl


# This function builds the MIME email structure
def MIMEmail(sender, to, subject, content):
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = to
    message['Subject'] = subject
    message['Reply-To'] = "TrippNT Sales Team <sales@trippnt.com>"
    message.attach(MIMEText(content, 'plain'))
    return message


with open(r'C:\Scripts\Stored Data\Security.csv', 'r') as secure:
    passwords = {}
    scsv = csv.reader(secure, delimiter=',')
    for row in scsv:
        if not row[1] == "":
            passwords[row[0]] = row[1]
    secure.close()


smtp_server = passwords['smtp_server']
port = passwords['port']  # For starttls
sender_email = passwords['sender_email']
password = passwords['password']
context = ssl.create_default_context()

processing = []
errors = {}
items = {}
eightFiftySixes = []
eightFiftySix = {}
send856List = ["VWR INTERNATIONAL"]
eightFiftySixHeader = ["Transaction Type", "Accounting ID", "Shipment ID", "SCAC", "Carrier Pro #",
                                   "Bill of Lading", "Scheduled Delivery", "Ship Date", "Ship To Name",
                                   "Ship To Address - Line One", "Ship To Address - Line Two", "Ship To City",
                                   "Ship To State",
                                   "Ship To Zip", "Ship To Country", "Ship to Address Code", "Ship Via", "Ship To Type",
                                   "Packaging Type", "Gross Weight", "Gross Weight UOM", "# of Cartons Shipped",
                                   "Carrier Trailer #", "Trailer Initial", "Ship From Name",
                                   "Ship From Address - Line One",
                                   "Ship From Address - Line Two", "Ship From City", "Ship From State", "Ship From Zip",
                                   "Ship From Country", "Ship From Address Code", "Vendor #", "DC Code",
                                   "Transportation Method", "Product Group", "Status", "Time Shipped", "PO #",
                                   "PO Date",
                                   "Invoice #", "Order Weight", "Store Name", "Store Number", "Mark For Code",
                                   "Department #",
                                   "Order Lading Quantity", "Packaging Type", "UCC-128", "Pack Size",
                                   "Inner Pack Per Outer Pack", "Pack Height", "Pack Length", "Pack Width",
                                   "Pack Weight",
                                   "Qty of UPCs within Pack", "UOM of UPCs", "Store Name", "Store Number", "Line #",
                                   "Vendor Part #", "Buyer Part #", "UPC #", "Item Description", "Quantity Shipped",
                                   "UOM",
                                   "Quantity Ordered", "Unit Price", "Pack Size", "Pack UOM",
                                   "Inner Packs per Outer Pack"]
# folder_path = 'H:\Order Entry\KARTS files\Sample 810s'
folder_path = 'C:\EDI-TempHome'
file_type = '\*csv'
files = glob.glob(folder_path + file_type)# Makes a list of all the csv files in the folder
counter2 = 0
filesAtStart = len(files)
if filesAtStart < 1:
    print("Couldn't find any files in the folder. Shutting down...")
    raise SystemExit
for file in files:
    processing.clear()
    file_creation_date = os.path.getctime(file)
    itemsOnOrder = []
    moveTo = "_Archive"
    with open(file, 'r', errors='replace') as invoice:
        rcsv = csv.reader(invoice, delimiter=',')
        for row in rcsv:
            if row[1] == "Accounting ID":
                headers = row[:]
            else:
                processing.append(row[:])
                invoiceNum = row[4]
                itemsOnOrder.append(row[36])
        for lineItem in processing:
            customer = lineItem[1]
            # Determines the customer type VWR, Fisher, and Global
            if customer == "VWR INTERNATIONAL":
                # Figures out VWR dropship based on where the shipment is going
                if 'VWR' in lineItem[7]:
                    if 'XDROP' in itemsOnOrder:
                        error = {
                            "Customer": customer,
                            "PO Number": lineItem[2],
                            "Error": "Dropship where there shouldn't be one"
                        }
                        errors[lineItem[4]] = error
                        moveTo = "_Error"
                else:
                    if 'XDROP' not in itemsOnOrder:
                        error = {
                            "Customer": customer,
                            "PO Number": lineItem[2],
                            "Error": "Missing Dropship Fee"
                        }
                        errors[lineItem[4]] = error
                        moveTo = "_Error"
            # Fisher needs a shipping line
            elif customer == "FISHER SCIENTIFIC":
                if 'SHIPPING' not in itemsOnOrder:
                    error = {
                        "Customer": customer,
                        "PO Number": lineItem[2],
                        "Error": "Missing Shipping Charge"
                    }
                    errors[lineItem[4]] = error
                    moveTo = "_Error"
            # Global needs tracking number
            elif customer == "GLOBAL":
                if len(lineItem[49]) <= 3:
                    error = {
                        "Customer": customer,
                        "PO Number": lineItem[2],
                        "Error": "Missing/Invalid Tracking Number"
                    }
                    errors[lineItem[4]] = error
                    moveTo = "_Error"
            # If not VWR/Fisher then make sure the order doesn't have shipping or dropship fees
            if 'SHIPPING' in itemsOnOrder and not customer == "FISHER SCIENTIFIC":
                error = {
                    "Customer": customer,
                    "PO Number": lineItem[2],
                    "Error": "Shipping charge where there shouldn't be one"
                }
                errors[lineItem[4]] = error
                moveTo = "_Error"
            elif 'XDROP' in itemsOnOrder and not customer == "VWR INTERNATIONAL":
                error = {
                    "Customer": customer,
                    "PO Number": lineItem[2],
                    "Error": "Dropship where there shouldn't be one"
                }
                errors[lineItem[4]] = error
                moveTo = "_Error"

    invoice.close()
    if moveTo == "_Archive":
        shutil.copy(file, r'C:\\True Commerce\\Transaction Manager\\Import\\810'
                    + datetime.fromtimestamp(file_creation_date).strftime("_%d_%H_%M_") + invoiceNum + ".csv")
        if customer in send856List:
            output2 = open(r'C:\\True Commerce\\Transaction Manager\\Import\\856'
                           + datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M_") + invoiceNum +
                           ".csv", 'w', newline='')
            writer2 = csv.writer(output2)
            writer2.writerow(eightFiftySixHeader)
            for row in processing:
                line = []
                eightFiftySix["Transaction Type"] = "856"
                eightFiftySix["Accounting ID"] = row[1]
                eightFiftySix["Shipment ID"] = row[4]
                eightFiftySix["SCAC"] = ""
                eightFiftySix["Carrier Pro #"] = row[49]
                eightFiftySix["Bill of Lading"] = ""
                eightFiftySix["Scheduled Delivery"] = ""
                eightFiftySix["Ship Date"] = row[24]
                eightFiftySix["Ship To Name"] = row[7]
                eightFiftySix["Ship To Address - Line One"] = row[8]
                eightFiftySix["Ship To Address - Line Two"] = row[9]
                eightFiftySix["Ship To City"] = row[10]
                eightFiftySix["Ship To State"] = row[11]
                eightFiftySix["Ship To Zip"] = row[12]
                eightFiftySix["Ship To Country"] = row[13]
                eightFiftySix["Ship to Address Code"] = ""
                eightFiftySix["Ship Via"] = row[23]
                eightFiftySix["Ship To Type"] = ""
                eightFiftySix["Packaging Type"] = ""
                eightFiftySix["Gross Weight"] = ""
                eightFiftySix["Gross Weight UOM"] = ""
                eightFiftySix["# of Cartons Shipped"] = ""
                eightFiftySix["Carrier Trailer #"] = ""
                eightFiftySix["Trailer Initial"] = ""
                eightFiftySix["Ship From Name"] = "TrippNT"
                eightFiftySix["Ship From Address - Line One"] = "10991 N Airworld Dr"
                eightFiftySix["Ship From Address - Line Two"] = ""
                eightFiftySix["Ship From City"] = "Kansas City"
                eightFiftySix["Ship From State"] = "MO"
                eightFiftySix["Ship From Zip"] = "64153"
                eightFiftySix["Ship From Country"] = "US"
                eightFiftySix["Ship From Address Code"] = "AIRWORLD"
                eightFiftySix["Vendor #"] = ""
                eightFiftySix["DC Code"] = ""
                eightFiftySix["Transportation Method"] = ""
                eightFiftySix["Product Group"] = ""
                eightFiftySix["Status"] = ""
                eightFiftySix["Time Shipped"] = ""
                eightFiftySix["PO #"] = row[2]
                eightFiftySix["PO Date"] = row[3]
                eightFiftySix["Invoice #"] = row[4]
                eightFiftySix["Order Weight"] = ""
                eightFiftySix["Store Name"] = ""
                eightFiftySix["Store Number"] = row[14]
                eightFiftySix["Mark For Code"] = ""
                eightFiftySix["Department #"] = ""
                eightFiftySix["Order Lading Quantity"] = ""
                eightFiftySix["Packaging Type"] = ""
                eightFiftySix["UCC-128"] = ""
                eightFiftySix["Pack Size"] = row[43]
                eightFiftySix["Inner Pack Per Outer Pack"] = row[44]
                eightFiftySix["Pack Height"] = ""
                eightFiftySix["Pack Length"] = ""
                eightFiftySix["Pack Width"] = ""
                eightFiftySix["Pack Weight"] = ""
                eightFiftySix["Qty of UPCs within Pack"] = ""
                eightFiftySix["UOM of UPCs"] = ""
                eightFiftySix["Store Name"] = ""
                eightFiftySix["Store Number"] = ""
                eightFiftySix["Line #"] = row[35]
                eightFiftySix["Vendor Part #"] = row[36]
                eightFiftySix["Buyer Part #"] = row[37]
                eightFiftySix["UPC #"] = row[38]
                eightFiftySix["Item Description"] = row[39]
                eightFiftySix["Quantity Shipped"] = row[40]
                eightFiftySix["UOM"] = "EA"
                eightFiftySix["Quantity Ordered"] = ""
                eightFiftySix["Unit Price"] = row[42]
                eightFiftySix["Pack Size"] = row[43]
                eightFiftySix["Pack UOM"] = "EA"
                eightFiftySix["Inner Packs per Outer Pack"] = row[44]
                for item in eightFiftySixHeader:
                    line.append(eightFiftySix[item])
                writer2.writerow(line)
            output2.close()
    shutil.move(file, folder_path + moveTo + "\\" + invoiceNum + ".csv")
if len(errors) > 0:
    message = MIMEmail("TrippNT DoNotReply <KARTS@trippnt.com>", "<AR@TrippNT.com>",
                       "KARTS detected errors in an" +
                       " outbound 810 file", "KARTS detected the following errors: \n\n" + str(errors))
    encoded_email = message.as_string()
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(message["From"], message["To"], encoded_email)

