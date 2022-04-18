
"""Import Statements------------------------------------------------------------------------------------------------"""
import csv
import glob
import os.path
import shutil
from datetime import *
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import ssl
import holidays
import requests
import time
import json
import math

"""Functions---------------------------------------------------------------------------------------------------------"""


def validateAddress(addressesToValidate):
    url = "https://apis-sandbox.fedex.com/address/v1/addresses/resolve"
    payload = {"addressesToValidate": list(addressesToValidate)}
    headers = {
        'Content-Type': "application/json",
        'X-locale': "en_US",
        'Authorization': token
        }
    response = requests.request("POST", url, data=json.dumps(payload).encode('utf8'), headers=headers)
    return response


# This function builds the MIME email structure
def MIMEmail(sender, to, subject, content):
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = to
    message['Bcc'] = "orders@trippnt.com"
    message['Subject'] = subject
    message['Reply-To'] = "TrippNT Sales Team <sales@trippnt.com>"
    message.attach(MIMEText(content, 'plain'))
    if "trippnt" not in str(to):
        message.attach(MIMEText(signatureHTML, "html"))
    return message


# checks to see if n is an integer
def is_integer(n):
    try:
        # will cast value n into an integer and see if it throws an error
        int(n)
    except ValueError:
        return False
    else:
        return True


# this function counts the number of characters in a string and returns the count of it.
def count_char(check, compare_to):
    count = 0
    for character in compare_to:
        if character == check:
            count += 1
    return count


# stores the customer emails for specific customers
def get_email(customercode, customeraddresses):
    if customercode == "GLOBAL":
        return customeraddresses['Global Main']['Primary Email']
    elif customercode == "FISHER SCIENTIFIC":
        return customeraddresses['Fisher Scientific']['Primary Email']
    else:
        # for anyone not Fisher or Global sends email back to TrippNT sales
        return "Sales@TrippNT.com"


# Generates email body based on rejection reason
def generate_message(contact_name, poNum, reason, partNum, price):
    message_text = ""
    if reason == "Pricing":
        message_text = ("Dear " + contact_name + ",\n\n" + "We are writing to inform you that PO " + poNum +
                        " has been rejected due to a price discrepancy. " +
                        "Please send in a revised PO with corrected pricing for your order to be entered and released" +
                        " to production.\n\n" + "The correct price for part number " + str(partNum) + " is $" +
                        str("{0:.2f}".format(price)) + " USD\n\n" +
                        "Please do not hesitate to contact us if you have any questions.\n\nSincerely,\n\nThe TrippNTeam")
    elif reason == "Discontinued":
        message_text = ("Dear " + contact_name + ",\n\n" + "We are writing to inform you that PO " + poNum +
                        " has been rejected due to the inclusion of obsolete product " + partNum +
                        ". Please send in a revised PO without said obsolete product for your order to be entered and " +
                        "sent to production.\n\nObsolete " + partNum + "'s comparable product is [product]. " +
                        "Please reference attached images. [comparablepartnumber] has an MSRP of &x.xx USD " +
                        "with your discount it will be $x.xx USD.\n\n" +
                        "Please do not hesitate to contact us if you have any questions.\n\nSincerely,\n\nThe TrippNTeam")
    return message_text


# calculates how many items of a set size can fit in a container (see FreightQuoter.py)
def howManyFit(item, qty, packageOptions, pad):
    selectedPackage = []
    qty1 = qty
    maxPerLayer = 0
    layerHeight = 0
    while len(selectedPackage) == 0 and qty1 > 0:
        for package in packageOptions:
            dir1 = math.floor((package[0] - pad) / item[0]) * math.floor((package[1] - pad) / item[1]) * math.floor(
                (package[2] - pad) / item[2])
            dir2 = math.floor((package[0] - pad) / item[1]) * math.floor((package[1] - pad) / item[2]) * math.floor(
                (package[2] - pad) / item[0])
            dir3 = math.floor((package[0] - pad) / item[2]) * math.floor((package[1] - pad) / item[0]) * math.floor(
                (package[2] - pad) / item[1])
            dir4 = math.floor((package[0] - pad) / item[0]) * math.floor((package[1] - pad) / item[2]) * math.floor(
                (package[2] - pad) / item[1])
            dir5 = math.floor((package[0] - pad) / item[1]) * math.floor((package[1] - pad) / item[0]) * math.floor(
                (package[2] - pad) / item[2])
            dir6 = math.floor((package[0] - pad) / item[2]) * math.floor((package[1] - pad) / item[1]) * math.floor(
                (package[2] - pad) / item[0])
            maxPerPack = max(dir1, dir2, dir3, dir4, dir5, dir6)
            if maxPerPack >= qty1:
                selectedPackage = package
                break
        if len(selectedPackage) == 0:
            qty1 = maxPerPack
    if qty1 > 0:
        lastPackQty = qty % qty1
    else:
        lastPackQty = 0
    if maxPerPack > qty:
        qtyPerPack = qty
    else:
        qtyPerPack = maxPerPack
    output = {"selectedPackage": selectedPackage,
              "maxPerPack": maxPerPack,
              "qtyPerPack": qtyPerPack,
              "lastPackQty": lastPackQty,
              }
    return output


# Converts old mapping to the new mapping
def convertToE2(lineItem, oMap, webOrder, posWithAlerts):
    nMap = {}
    custName = lineItem[oMap["ACCOUNTING ID"]]
    shipMethod = lineItem[oMap["SHIP VIA"]].upper()
    nMap["TRANSACTION ID"] = lineItem[oMap["TRANSACTION ID"]]
    nMap["ACCOUNTING ID"] = lineItem[oMap["ACCOUNTING ID"]]
    nMap["PURCHASE ORDER NUMBER"] = lineItem[oMap["PURCHASE ORDER NUMBER"]] + "-NP"
    if not lineItem[oMap["PURCHASE ORDER NUMBER"]] in posWithAlerts:
        if custName == "AMAZON" or custName == "WAYFAIR":
            nMap["PURCHASE ORDER NUMBER"] = lineItem[oMap["PURCHASE ORDER NUMBER"]]
        elif (custName == "GLOBAL" or custName == "VWR INTERNATIONAL") and (shipMethod == "UPS" or
                                                                            shipMethod == "UPS GROUND"):
            nMap["PURCHASE ORDER NUMBER"] = lineItem[oMap["PURCHASE ORDER NUMBER"]]
        elif custName == "FISHER SCIENTIFIC" and shipMethod == "TRIPPNT GROUND":
            nMap["PURCHASE ORDER NUMBER"] = lineItem[oMap["PURCHASE ORDER NUMBER"]]
    nMap["PO DATE"] = lineItem[oMap["PO DATE"]]
    nMap["SHIP TO NAME"] = lineItem[oMap["SHIP TO NAME"]]
    nMap["SHIP TO ADDRESS 1"] = lineItem[oMap["SHIP TO ADDRESS 1"]]
    nMap["SHIP TO ADDRESS 2"] = lineItem[oMap["SHIP TO ADDRESS 2"]]
    nMap["SHIP TO CITY"] = lineItem[oMap["SHIP TO CITY"]]
    nMap["SHIP TO STATE"] = lineItem[oMap["SHIP TO STATE"]]
    nMap["SHIP TO ZIP"] = lineItem[oMap["SHIP TO ZIP"]]
    nMap["SHIP TO COUNTRY"] = lineItem[oMap["SHIP TO COUNTRY"]]
    nMap["STORE NUMBER"] = lineItem[oMap["STORE NUMBER"]]
    nMap["BILL TO NAME"] = lineItem[oMap["BILL TO NAME"]]
    nMap["BILL TO ADDRESS 1"] = lineItem[oMap["BILL TO ADDRESS 1"]]
    nMap["BILL TO ADDRESS 2"] = lineItem[oMap["BILL TO ADDRESS 2"]]
    nMap["BILL TO CITY"] = lineItem[oMap["BILL TO CITY"]]
    nMap["BILL TO STATE"] = lineItem[oMap["BILL TO STATE"]]
    nMap["BILL TO ZIP"] = lineItem[oMap["BILL TO ZIP"]]
    nMap["BILL TO COUNTRY"] = lineItem[oMap["BILL TO COUNTRY"]]
    nMap["BILL TO CODE"] = lineItem[oMap["BILL TO CODE"]]
    if lineItem[oMap["SHIP VIA"]].replace("UPS", "") == "":
        lineItem[oMap["SHIP VIA"]] = "UPS GROUND"
    commaIndex = lineItem[oMap["SHIP VIA"]].find(",")
    if commaIndex > 0:
        nMap["SHIP VIA"] = lineItem[oMap["SHIP VIA"]][0:commaIndex]
    else:
        nMap["SHIP VIA"] = lineItem[oMap["SHIP VIA"]]
    if len(nMap["SHIP VIA"]) > 20:
        nMap["SHIP VIA"] = nMap["SHIP VIA"][0:19]
    nMap["SHIP DATE"] = lineItem[oMap["SHIP DATE"]]
    nMap["TERMS"] = lineItem[oMap["TERMS"]]
    nMap["NOTE"] = lineItem[oMap["NOTE"]]
    nMap["DEPARTMENT NUMBER"] = lineItem[oMap["DEPARTMENT NUMBER"]]
    nMap["CANCEL DATE"] = lineItem[oMap["CANCEL DATE"]]
    nMap["DO NOT SHIP BEFORE"] = lineItem[oMap["DO NOT SHIP BEFORE"]]
    nMap["DO NOT SHIP AFTER"] = lineItem[oMap["DO NOT SHIP AFTER"]]
    nMap["Allowance Percent1"] = lineItem[oMap["Allowance Percent1"]]
    nMap["Allowance Amount1"] = lineItem[oMap["Allowance Amount1"]]
    nMap["Allowance Precent2"] = lineItem[oMap["Allowance Precent2"]]
    nMap["Allowance Amount2"] = lineItem[oMap["Allowance Amount2"]]
    nMap["LINE #"] = lineItem[oMap["LINE #"]]
    nMap["VENDOR PART #"] = lineItem[oMap["VENDOR PART #"]]
    nMap["BUYERS PART #"] = lineItem[oMap["BUYERS PART #"]]
    nMap["UPC #"] = lineItem[oMap["UPC #"]]
    nMap["DESCRIPTION"] = lineItem[oMap["DESCRIPTION"]]
    nMap["QUANTITY"] = lineItem[oMap["QUANTITY"]]
    nMap["UOM"] = lineItem[oMap["UOM"]]
    nMap["UNIT PRICE"] = lineItem[oMap["UNIT PRICE"]]
    nMap["ITEM NOTES"] = lineItem[oMap["ITEM NOTES"]]
    nMap["CUSTOMER ORDER #"] = lineItem[oMap["CUSTOMER ORDER #"]]
    nMap["BUYER NAME"] = lineItem[oMap["BUYER NAME"]]
    nMap["BUYER PHONE"] = lineItem[oMap["BUYER PHONE"]]
    nMap["BUYER FAX"] = lineItem[oMap["BUYER FAX"]]
    nMap["Placeholder"] = str(lineItem[oMap["SHIP TO NAME"]])
    nMap["BUYER EMAIL"] = lineItem[oMap["BUYER EMAIL"]]
    nMap["INFO CONTACT NAME"] = lineItem[oMap["INFO CONTACT NAME"]]
    nMap["INFO CONTACT PHONE"] = lineItem[oMap["INFO CONTACT PHONE"]]
    nMap["INFO CONTACT EMAIL"] = lineItem[oMap["INFO CONTACT EMAIL"]]
    nMap["DEL CONTACT NAME"] = lineItem[oMap["DEL CONTACT NAME"]]
    nMap["DEL CONTACT PHONE"] = lineItem[oMap["DEL CONTACT PHONE"]]
    nMap["DEL CONTACT EMAIL"] = lineItem[oMap["DEL CONTACT EMAIL"]]
    nMap["DEL REFERENCE"] = lineItem[oMap["DEL REFERENCE"]]
    nMap["SHIP TO ADDRESS 3"] = lineItem[oMap["SHIP TO ADDRESS 3"]]
    nMap["Item Notes"] = lineItem[oMap["Item Notes"]]
    nMap["Ship_Description"] = lineItem[oMap["SHIP VIA"]]
    nMap["Company_Code"] = ""
    nMap["EDI_Order_Number"] = ""
    nMap["EDI_Item_Number"] = ""
    nMap["Salesman_Code"] = ""
    nMap["Terms_Code"] = ""
    nMap["Tax_Code"] = ""
    nMap["GST_Code"] = ""
    nMap["Exchange_Rate"] = ""
    nMap["Order_Header_Notes"] = ""
    nMap["Shipping_Address_Location"] = ""
    nMap["Shipping_Address_Phone_Number"] = lineItem[oMap["DEL CONTACT PHONE"]]
    if commaIndex > 0:
        poundIndex = lineItem[oMap["SHIP VIA"]].find("#")
        acctIndex = lineItem[oMap["SHIP VIA"]].upper().find("ACCT")
        if poundIndex > 0:
            nMap["Shipping_Account_Number"] = lineItem[oMap["SHIP VIA"]][poundIndex:]
        elif acctIndex > 0:
            nMap["Shipping_Account_Number"] = lineItem[oMap["SHIP VIA"]][acctIndex:]
        else:
            nMap["Shipping_Account_Number"] = lineItem[oMap["SHIP VIA"]][commaIndex:]
    else:
        nMap["Shipping_Account_Number"] = ""
    nMap["Billing_Address_Location"] = lineItem[oMap["BILL TO NAME"]]
    nMap["Header_User_Date1"] = ""
    nMap["Header_User_Date2"] = ""
    nMap["Header_User_Date1"] = ""
    nMap["Header_User_Date2"] = ""
    timestamp = lineItem[oMap["ISA Timestamp"]].split(" ")
    dated = datetime.strptime(timestamp[0], '%m/%d/%Y')
    try:
        timed = timestamp[1]
        timed = timed.split(":")
        if timestamp[2] == "PM":
            if int(timed[0]) < 12:
                timed[0] = str(int(timed[0]) + 12)
        else:
            if int(timed[0]) == 12:
                timed[0] = "00"
        timed[0] = "00" + timed[0]
        timed[1] = "00" + timed[1]
        timed[2] = "00" + timed[2]
        timed = timed[0][-2:]+":"+timed[1][-2:]+":"+timed[2][-2:]
        timestamp = datetime.combine(dated, datetime.min.time())
        nMap["Header_User_Text1"] = timestamp.strftime("%m/%d/%Y")+" "+timed
    except IndexError:
        print("Couldn't parse time")
        nMap["Header_User_Text1"] = lineItem[oMap["ISA Timestamp"]]
        timestamp = datetime.combine(dated, datetime.min.time())
    if len(lineItem[oMap["DO NOT SHIP BEFORE"]]) > 0:
        doNotShipUntil = datetime.strptime(lineItem[oMap["DO NOT SHIP BEFORE"]], '%m/%d/%Y')
        doNotShipUntil = datetime.combine(doNotShipUntil, datetime.min.time())
        if timestamp < doNotShipUntil:
            timestamp = doNotShipUntil + timedelta(seconds=1)
            nMap["Header_User_Text1"] = timestamp.strftime("%m/%d/%Y %H:%M:%S")
    if webOrder:
        nMap["Header_User_Text2"] = "RETAIL"
    else:
        nMap["Header_User_Text2"] = ""
    nMap["Header_User_Text3"] = ""
    nMap["Header_User_Text4"] = ""
    nMap["Header_User_Currency1"] = ""
    nMap["Header_User_Currency2"] = ""
    nMap["Header_User_Number1"] = ""
    nMap["Header_User_Number2"] = ""
    nMap["Header_User_Number3"] = ""
    nMap["Header_User_Number4"] = ""
    nMap["Header_User_Memo1"] = ""
    nMap["Status"] = "Automatic"
    nMap["Discount_Percent"] = ""
    nMap["Miscellaneous_Charge1"] = ""
    nMap["Miscellaneous_Description1"] = ""
    nMap["Rate_Code"] = ""
    nMap["Work_Code"] = ""
    nMap["Product_Code"] = ""
    nMap["Priority"] = ""
    nMap["FOB"] = ""
    nMap["Commission_Percent"] = ""
    nMap["Order_Detail_Notes"] = ""
    nMap["Detail_User_Date1"] = ""
    nMap["Detail_User_Date2"] = ""
    nMap["Detail_User_Text1"] = ""
    nMap["Detail_User_Text2"] = ""
    nMap["Detail_User_Text3"] = ""
    nMap["Detail_User_Text4"] = ""
    nMap["Detail_User_Currency1"] = ""
    nMap["Detail_User_Currency2"] = ""
    nMap["Detail_User_Number1"] = ""
    nMap["Detail_User_Number2"] = ""
    nMap["Detail_User_Number3"] = ""
    nMap["Detail_User_Number4"] = ""
    nMap["Detail_User_Memo1"] = ""
    nMap["Order_Release_Due_Date"] = ""
    nMap["Order_Release_Notes"] = ""
    nMap["Miscellaneous_Charge_Code"] = ""
    nMap["Miscellaneous_Charge_Amount"] = ""
    nMap["Miscellaneous_Charge_Include_In_Piece_Price"] = ""
    nMap["Miscellaneous_Charge_Description"] = ""
    nMap["Order_Type"] = "Manufacturing"
    nMap["Order_Detail_Freight_Amount"] = ""
    nMap["Division_Number"] = ""
    nMap["Order_Release_Release_Number"] = ""
    nMap["Residential"] = ""
    nMap["Print_Certification"] = ""
    nMap["Miscellaneous_Charge_Unit_Cost"] = ""
    nMap["Miscellaneous_Charge_Unit_Of_Measure"] = ""
    nMap["EDI_Status"] = "Automatic"
    nMap["Exported"] = ""
    nMap["Lock_Price"] = ""
    nMap["Trading_Partner"] = ""
    nMap["Address_Location_Number"] = ""
    return nMap.values()


def createSQL(poHeader, freight, sqlCounter):
    # brings in stored customer freight remittance addresses
    with open(r'C:\Scripts\Stored Data\CustomerAddresses.csv', 'r') as customerAddresses:
        remitAddresses = {}
        csvreader = csv.reader(customerAddresses, delimiter=',')
        for row in csvreader:
            if row[0] == "KARTS Address Code":
                addressCode = row[1]
                remitAddresses[addressCode] = {}
            else:
                remitAddresses[addressCode][row[0]] = row[1]
        customerAddresses.close()
    # sets default values for all fields
    customer = poHeader[1]
    documentDate = datetime.today().strftime("%d%H")
    vendorCodeFreight = "FEDEX"
    vendorNameFreight = "FEDEX"
    shippingService = 'FEDEX GROUND'
    shipDescription = 'FedEx Ground'
    shipCode = 'TrippNT Ground'
    shippingChargePrepaid = 1
    shippingChargePaidBy = 'Sender'
    addShippingChargesToInvoice = 1
    addressCode = 'Airworld'
    termsCodeFreight = ''
    specialInstructions = ''
    freightCollectAccountNumber = ''
    customerPONumber = poHeader[2]
    apContactName = poHeader[12]
    apContactEmail = ''
    apContactPhone = ''
    contactEmail = poHeader[49]
    acknowledgedBy = 'KARTS'
    # applies logic (based on the customers' routing instructions) to select the correct address code & ship method
    if customer == "GLOBAL":
        commaIndex = poHeader[20].find(",")
        poundIndex = poHeader[20].find("#")
        punctuationIndex = max(commaIndex, poundIndex)
        shippingChargePrepaid = 0
        shippingChargePaidBy = 'Third Party'
        addShippingChargesToInvoice = 0
        addressCode = 'Global Main'
        apContactName = 'Accounts Payable'
        apContactEmail = 'hmuhammad@globalindustrial.com'
        apContactPhone = '(888) 978-7759'
        if len(poHeader[20].replace("UPS GROUND", "")) < len(poHeader[20]):
            vendorCodeFreight = 'UPS'
            vendorNameFreight = 'UPS'
            shippingService = 'UPS GROUND'
            shipDescription = poHeader[20]
            shipCode = 'UPS GROUND'
            freightCollectAccountNumber = '394310'
            if len(poHeader[20].replace("UPS GROUND", "")) > 0:
                freightCollectAccountNumber = poHeader[20][punctuationIndex+1:]
        elif len(poHeader[20].replace("FEDEX GROUND", "")) < len(poHeader[20]):
            vendorCodeFreight = 'FEDEX'
            vendorNameFreight = 'FEDEX'
            shippingService = 'FEDEX GROUND'
            shipDescription = poHeader[20]
            shipCode = 'FEDEX GROUND'
            freightCollectAccountNumber = poHeader[20][punctuationIndex+1:]
        elif len(poHeader[20].replace("NFI", "")) < len(poHeader[20]):
            vendorCodeFreight = 'NFI'
            vendorNameFreight = 'NFI TRUCKING'
            shippingService = 'NFI TRUCKING'
            shipDescription = poHeader[20]
            shipCode = 'LTL'
            addressCode = 'Global NFI'
        else:
            vendorCodeFreight = poHeader[20][0:commaIndex]
            vendorNameFreight = poHeader[20][0:commaIndex]
            shippingService = poHeader[20][0:commaIndex]
            shipDescription = poHeader[20]
            shipCode = 'LTL'
            acknowledgedBy = 'REVIEW'
    elif customer == "VWR INTERNATIONAL":
        if freight:
            vendorCodeFreight = "UPS FREIGHT"
            vendorNameFreight = "UPS FREIGHT"
            shippingService = 'UPS FREIGHT'
            shipDescription = 'Freight Shipment (UPS)'
            shipCode = 'LTL'
            addressCode = 'VWR Freight'
        else:
            vendorCodeFreight = "UPS"
            vendorNameFreight = "UPS"
            shippingService = 'UPS GROUND'
            shipDescription = 'UPS Ground'
            shipCode = 'UPS'
            freightCollectAccountNumber = '305R3A'
            addressCode = 'VWR Main'
            if "CA" in poHeader[10].upper():
                addressCode = 'VWR Canada'
                if "VWR" in poHeader[4].upper():
                    freightCollectAccountNumber = '4F4332'
                else:
                    freightCollectAccountNumber = '506737'
        shippingChargePrepaid = 0
        shippingChargePaidBy = 'Third Party'
        addShippingChargesToInvoice = 0
        apContactName = 'Accounts Payable'
        apContactPhone = '(816) 792-2604'
        apContactEmail = 'Sales@TrippNT.com'
    elif customer == "FISHER SCIENTIFIC":
        if freight:
            vendorCodeFreight = "FEDEX FREIGHT"
            vendorNameFreight = "FEDEX FREIGHT"
            shippingService = 'FEDEX FREIGHT ECONOMY'
            shipDescription = 'FedEx Freight'
            shipCode = 'TrippNT Freight'
        else:
            vendorCodeFreight = "FEDEX"
            vendorNameFreight = "FEDEX"
            shippingService = 'FEDEX GROUND'
            shipDescription = 'FedEx Ground'
            shipCode = 'TrippNT Ground'
            freightCollectAccountNumber = '289152784'
        shippingChargePrepaid = 1
        shippingChargePaidBy = 'Sender'
        addShippingChargesToInvoice = 1
        addressCode = 'Airworld'
        contactEmail = 'susan.fisher@thermofisher.com'
        apContactName = ''
        apContactPhone = '(816) 792-2604'
        apContactEmail = 'Sales@TrippNT.com'
    elif customer == "AMAZ*" or customer == "WAYFAIR":
        vendorCodeFreight = ''
        vendorNameFreight = ''
        shippingService = customer+' SHIPPING'
        shipDescription = 'SHIP VIA '+customer+' VENDOR CENTRAL'
        shipCode = customer+' SHIPPING'
        shippingChargePrepaid = 0
        if customer == "AMAZON":
            shippingChargePaidBy = 'Freight Collect'
        else:
            shippingChargePaidBy = 'Third Party'
        if customer == "WAYFAIR":
            specialInstructions = 'Please use customer packing slip'
        addShippingChargesToInvoice = 0
        frName = poHeader[12]
        frAddress = poHeader[13]+", "+poHeader[14]
        frCity = poHeader[15]
        frState = poHeader[16]
        frZip = poHeader[17]
        frCountry = poHeader[10]
        frAddressType = 'Customer'
        addressCode = 'SKIP'
        termsCodeFreight = ''
        specialInstructions = ''
        freightCollectAccountNumber = ''
        customerPONumber = poHeader[2]
        apContactName = poHeader[12]
        contactEmail = poHeader[49]
    # uses the selected address code to retrieve and set the address values
    if not addressCode == 'SKIP':
        frName = remitAddresses[addressCode]['Address Name']
        frAddress = remitAddresses[addressCode]['Address 1']+remitAddresses[addressCode]['Address 2']
        frCity = remitAddresses[addressCode]['City']
        frState = remitAddresses[addressCode]['State']
        frZip = remitAddresses[addressCode]['Zip']
        frCountry = remitAddresses[addressCode]['Country']
        frAddressType = remitAddresses[addressCode]['Address Type']
    # fills in the SQL statement with the values chosen above
    baseSQL = f"""/****** Script for UpdateOrderStuffs ******/
        UPDATE [TrippNT].[dbo].[Order_Header]
        SET
        [Vendor_Code_Freight] = '{vendorCodeFreight}',
        [Vendor_Name_Freight] = '{vendorNameFreight}',
        [Terms_Code_Freight] = '{termsCodeFreight}',
        [Special_Instructions] = '{specialInstructions}',
        [Shipping_Charge_Prepaid] = {shippingChargePrepaid},
        [Shipping_Charges_Paid_By] = '{shippingChargePaidBy}',
        [Shipping_Service] = '{shippingService}',
        [Notification] = 1,
        [Delivery] = 1,
        [Exception] = 1,
        [Notification_Email_Address] = '{contactEmail}',
        [Delivery_Email_Address] = '{contactEmail}',
        [Exception_Email_Address] = '{contactEmail}',
        [Freight_Collect_Account_Number] = '{freightCollectAccountNumber}',
        [Add_Shipping_Charges_To_Invoice] = {addShippingChargesToInvoice},
        [Employee_Code_Acknowledged_By] = '{acknowledgedBy}'

        WHERE Customer_PO_Number LIKE '{customerPONumber}%' AND Customer_Code = '{customer}'

        DECLARE @ohID{sqlCounter} int
        SELECT @ohID{sqlCounter} = [Order_Header_ID]
        FROM [TrippNT].[dbo].[Order_Header]
        WHERE Customer_PO_Number LIKE '{customerPONumber}%' AND Customer_Code = '{customer}'

        UPDATE [TrippNT].[dbo].[Address]
        SET
        [Name] = '{frName}',
        [Street_Address] = '{frAddress}',
        [City] = '{frCity}',
        [State_Code] = '{frState}',
        [Postal_Code] = '{frZip}',
        [Country_Code] = '{frCountry}',
        [Shipping_Description] = '{shipDescription}',
        [Shipping_Code] = '{shipCode}',
        [Contact_Name] = '{apContactName}',
        [Phone_Number] = '{apContactPhone}',
        [Email_Address] = '{apContactEmail}',
        [Source] = '{frAddressType}',
        [Location_Code] = ''
        WHERE [Order_Header_ID] = @ohID{sqlCounter} AND [Address_Type] = 'REMITTANCE'

        """
    mode = 'a'
    if sqlCounter <= 1:
        mode = 'w'
    sqlOutput = open('H:\\Order Entry\\KARTS files\\SQL\\sqlOutput'+str(documentDate)+'.txt', mode)
    sqlOutput.write(baseSQL)
    sqlOutput.close()


"""Variables---------------------------------------------------------------------------------------------------------"""

# Testing Mode Switch
testMode = False

with open(r'C:\Scripts\Stored Data\Security.csv', 'r') as secure:
    passwords = {}
    csvreader = csv.reader(secure)
    for row in csvreader:
        passwords[row[0]] = row[1]
    secure.close()

# Related to emailer
smtp_server = passwords['smtp_server']
port = passwords['port']  # For starttls
sender_email = passwords['sender_email']
password = passwords['password']
context = ssl.create_default_context()
signatureHTML = """ <html> <body> <table style="cell-padding:0px;boarder-spacing:0px;white-space: nowrap;"> <tbody> <tr> 
<td style="border-right:solid;border-right-width:3px;border-color:rgb(114, 180, 49);padding-right:10px" rowspan="2"> 
<a href="http://www.trippnt.com/Default.asp" rel="noopener noreferrer"> <img class="front-inline-attachment" 
src="https://static1.squarespace.com/static/5f31547a6c5578025c50408c/t/6100198f73ba7b738a90735d/1627396495457
/KartsPic.png" style="width: 140px; height: 140px;" width="140" height="140" front-cid="d5a5fe2d"> </a> <br> </td> 
<td style="padding-left: 10px; font-size: 14px; color: rgb(0, 76, 151);"> <span style="font-weight: bold; color: rgb(
114, 180, 49); font-size: 18px; line-height: 18.4px;">The TrippNTeam</span> <br> <span style="font-weight: bold; 
font-size: 16px;">KARTS Automated System</span> <br> <b>1.800.874.7768&nbsp;| </b> <a style="color: rgb(114, 180, 
49); white-space: nowrap;" rel="noopener noreferrer" href="mailto:Sales@TrippNT.com"> <b>Sales@TrippNT.com</b> </a> 
<br> <span style="font-weight: bold;">10991 N Airworld Drive, Kansas City, MO 64153</span> </td> </tr> <tr> <td 
style="padding-left:10px;padding-top:5px"> <a href="http://www.trippnt.com/Default.asp" rel="noopener noreferrer"> 
<img style="width: 140px; height: 52.2957px;" alt="TrippNT" class="front-inline-attachment" height="52" width="140" 
src="https://trippnt.com/v/images/TrippNTlogo.png" front-cid="58894f5c"> </a> </td> </tr> </tbody> </table> </body> 
</html> """

# Relating to the reject emails
rejected = []
portal_rejected = []
reject_emails = []
subject_reasons = []

# Customer contact type (email by default, if not set here.) Can be 855 and portal or email, but NOT portal AND email.
portalRejectList = ["AMAZON", "AMAZON DIRECT", "WAYFAIR"]
send855List = ["VWR INTERNATIONAL", "FISHER SCIENTIFIC", "AMAZON"]
fixPrice = ["WAYFAIR"]

# Relating to pricing
price_dict = {}
amzn_price_dict = {}

# Relating to PO processing
x = slice(0, 59)  # this slice is the columns from the orignal PO that will be written to the file
PO_Price = ()
PO_Number = ()
partNum = ()
processing = []
rejectReason = []
total_dict = {}
orderDimsDict = {}
freight = False
kartsNotes = {}

# Related to mapping
tcIndex = {}
e2Index = {}

# Related to flat file upload for TrueCommerce
POAUpload = []

# Lead Time/Backorder -- note that DAYS is not WORKDAYS. Weekends/holidays will be moved to the next available date.
cartLeadTime = timedelta(weeks=3)
fabLeadTime = timedelta(weeks=2)
kanbanLeadTime = timedelta(weeks=2)
kbanCartLeadTime = timedelta(days=2)
# Capitalize Each Word in this list. Items should be in order of LATEST to SOONEST.
backorderList = ["Example"]
# CAPITALIZE WHOLE STRING in this dictionary
backorderDict = {"EXAMPLE": datetime(2021, 12, 10, 0, 0)}

# Date and time (used for calculating ship dates, mostly)
todays_date = datetime.today()
thisYear = int(todays_date.strftime("%Y"))
nextYear = thisYear + 1
us_holidays = holidays.US()
trippNTHolidays = ["New Year's Day", "New Year's Day (Observed)", 'Memorial Day',
                   'Juneteenth National Independence Day',
                   'Independence Day', 'Independence Day (Observed)', 'Labor Day', 'Thanksgiving', 'Christmas Day',
                   'Christmas Day (Observed)']

# Buyer notes
notesFields = [23, 40, 55]

"""SETUP BEGINS HERE ------------------------------------------------------------------------------------------------"""

# makes a list of TrippNT days off (aka "holiDates")
holiDates = []
for ptr in holidays.US(years=thisYear).items():
    # adds the TrippNT Holidays to the list of days off
    # ptr pulls the date and corresponding holiday index 0 is the day and index 1 is the holiday name
    if ptr[1] in trippNTHolidays:
        holiDates.append(ptr[0])
    # adds Black Friday
    if ptr[1] == 'Thanksgiving':
        blackFriday = ptr[0] + timedelta(days=1)
        holiDates.append(blackFriday)
# adds the days between Christmas and New Year's
xmasCounter = 26
while xmasCounter <= 31:
    holiDates.append(date(thisYear, 12, xmasCounter))
    xmasCounter += 1
holiDates.append(date(nextYear, 1, 1))

# FedEx credential request
url = "https://apis-sandbox.fedex.com/oauth/token"

payload = {'grant_type': 'client_credentials', 'client_id': passwords['clientID'],
           'client_secret': passwords['clientSecret']}
headers = {
    'Content-Type': "application/x-www-form-urlencoded"
    }
response = requests.request("POST", url, data=payload, headers=headers)

token = "Bearer "+response.json()["access_token"]

# Reads in Kanban list
with open(r'C:\Scripts\Stored Data\Kanban Items.csv', 'r') as kanban:
    kanbanList = []
    for row in kanban:
        kanbanList.append(row[0:5])
    kanban.close()

# Reads in product pricing
price_path = r'C:\Scripts\Stored Data\Price_Sheet.csv'
counter = 0
with open(price_path, 'r') as price_sheet:
    pcsv = csv.reader(price_sheet, delimiter=',')
    for row in pcsv:
        if counter > 0:
            part_num = row[0]
            price = float(row[3])
            price_dict[part_num] = price
        counter += 1
        price_dict["SHIPPING"] = 0
        price_dict["DISCOUNT"] = 0
        price_dict["00000"] = 0
        price_dict["10000"] = 0
        price_dict["SHIPP"] = 0
        price_dict["DISCO"] = 0
    price_sheet.close()

# Reads in Amazon's special pricing
amznPricePath = r'C:\Scripts\Stored Data\AmazonPricing.csv'
counter = 0
with open(amznPricePath, 'r') as amznPriceSheet:
    pcsv = csv.reader(amznPriceSheet, delimiter=',')
    for row in pcsv:
        if counter > 0:
            part_num = row[0]
            price = float(row[3])
            amzn_price_dict[part_num] = price
        counter += 1
    amznPriceSheet.close()

# Creates a dictionary of the header so that columns can be called by name instead of index
with open(r'C:\Scripts\Stored Data\e2.csv', 'r') as e2Header:
    headerReader = csv.reader(e2Header, delimiter=',')
    columnCounter = 0
    for row in headerReader:
        for item in row:
            e2Index[item] = columnCounter
            columnCounter += 1
    e2Header.close()

pad = 2  # amount of padding, in inches, needed on each side of the box.
palletPad = 0.5  # amount of padding, in inches, needed on each side of the pallet.
maxGroundWeight = 150
maxGroundCube = 60
with open(r'C:\Scripts\Stored Data\Weights_n_Dims.csv', 'r') as weightsNdims:
    weightsNdimsDict = {}
    rcsv = csv.reader(weightsNdims, delimiter=',')
    counter = 0
    for row in rcsv:
        if counter > 0:
            weightsNdimsDict[row[0]] = {"width": float(row[1]), "height": float(row[2]), "depth": float(row[3]),
                                        "weight": float(row[4]), "ltlThreshold": row[5]}
        counter += 1
    weightsNdims.close()
boxes = [[12, 12, 6], [12, 12, 12], [14, 14, 14], [20, 14, 12], [26, 14, 16], [18, 18, 18], [27, 13, 23], [21, 21, 21],
         [24, 24, 24]]
pallets = [[30, 50, 66, 30], [40, 48, 66, 35]]

"""CORE FUNCTIONALITY BEGINS HERE------------------------------------------------------------------------------------"""
if testMode:
    folder_path = r'C:\\Users\\kheimonen\\Desktop\\Order Entry Stuff\\Data Processing'
else:
    folder_path = r'C:\\True Commerce\\Transaction Manager\\Export'
file_type = '\\*csv'
files = glob.glob(folder_path + file_type)  # Makes a list of all the csv files in the folder
# Checks to make sure the list is not empty
if len(files) <= 0:
    print("Unable to find file. Shutting down...")
    raise SystemExit
maxFile = max(files, key=os.path.getctime)  # Chooses the newest file in the folder
fileName = str(os.path.basename(maxFile))
if fileName.find("Web") >= 0:
    webOrder = True
else:
    webOrder = False
# generates an output file called "Purchase_Orderdd_hh_mm" (where date is based on the creation date of maxFile)
file_creation_date = os.path.getctime(maxFile)
fileEnding = datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M")
if webOrder:
    fileEnding = fileEnding + "W"
if testMode:
    output = open("H:\\Order Entry\\EDI 2\\Purchase_Order_K_" + fileEnding + ".csv", 'w', newline='')
else:
    output = open("\\\\sql\\Users\\tntadmin\\Desktop\\Import\\Purchase_Order_K_" + fileEnding + ".csv", 'w', newline='')
writer = csv.writer(output)

# Reads import file into processing
with open(maxFile, 'r', errors='replace') as PO:
    rcsv = csv.reader(PO, delimiter=',')
    counter = 0
    for row in rcsv:
        if counter > 0:
            shipState = row[8]
            shipZip = row[9]
            shipCountry = row[10]
            if len(row[33]) > 5:
                if row[33] == "SHIPPING":
                    row[33] = '00000'
                else:
                    partNum = row[33][0:5]
                    row[33] = partNum
            else:
                partNum = row[33]
            PO_Number = row[2]
            PO_Price = float(row[39])  # this could, theoretically, cause a type conversion error (but so far it hasn't)
            customer = row[1]
            # Adds hyphens to long format zip codes
            if len(row[9]) == 9:
                row[9] = row[9][:5] + "-" + row[9][-4:]
            # Calculates ship date based on PO date, customer, kanban or not, and current date
            # no ship date will fill with today's date
            if row[3] == "":
                po_date = todays_date
            else:
                po_date = datetime.strptime(row[3], '%m/%d/%Y')
            if not (row[1] == "AMAZON DIRECT" or row[1] == "WAYFAIR" or row[1] == "AMAZON"):
                # no ship date will fill with today's date
                # Checks to see if part number is in Kanban list and applies proper lead time
                if row[33] in kanbanList:
                    if "CART" in row[36].upper():
                        shipDate = po_date + kbanCartLeadTime
                    else:
                        shipDate = po_date + kanbanLeadTime
                elif "CART" in row[36].upper():
                    shipDate = po_date + cartLeadTime
                else:
                    shipDate = po_date + fabLeadTime
                # makes sure shipDate is formatted the same as holiDate(s)
                shipDate = date(shipDate.year, shipDate.month, shipDate.day)
                # loops until it finds a weekday that is not a holiDate (maybe reformat on later date)
                counter = 0
                while counter < 1:
                    if shipDate in holiDates:
                        shipDate += timedelta(days=1)
                    elif shipDate.weekday() == 5:
                        shipDate += timedelta(days=2)
                    elif shipDate.weekday() == 6:
                        shipDate += timedelta(days=1)
                    else:
                        counter += 1
                # Conversion back to datetime object (prevents comparison errors)
                shipDate = datetime.combine(shipDate, datetime.min.time())
                # sets past ship days to the current day
                if todays_date > shipDate:
                    shipDate = todays_date
                for product in backorderList:
                    if product in row[36] or product.upper() in row[36] or product.lower() in row[36]:
                        backorderDate = backorderDict[product.upper()]
                        if backorderDate > shipDate:
                            shipDate = backorderDate
                # retrieves the ship date from the PO and compares it to our shipdate. Uses whichever is later.
                if row[21] == "":
                    po_shipDate = shipDate
                else:
                    po_shipDate = datetime.strptime(row[21], '%m/%d/%Y')
                if not po_shipDate > shipDate:
                    deliverDate = shipDate + timedelta(days=3)
                    if deliverDate.weekday() == 5:
                        deliverDate += timedelta(days=2)
                    elif shipDate.weekday() == 6:
                        deliverDate += timedelta(days=1)
                    deliverDate = deliverDate.strftime("%m/%d/%Y")
                    shipDate = shipDate.strftime("%m/%d/%Y")
                    row[21] = shipDate
                else:
                    deliverDate = po_shipDate + timedelta(days=3)
                    if deliverDate.weekday() == 5:
                        deliverDate += timedelta(days=2)
                    elif shipDate.weekday() == 6:
                        deliverDate += timedelta(days=1)
                    deliverDate = deliverDate.strftime("%m/%d/%Y")
            else:
                try:
                    shipDate = datetime.strptime(row[21], '%m/%d/%Y')
                except:
                    shipDate = row[27]
            # if ship date is blank then fill it with "do not ship until" date
            if row[21] == "":
                row[21] = row[27]
            # Sets shipping codes to their default values
            if row[1] == "AMAZON" or row[1] == "AMAZON DIRECT":
                row[20] = "AMAZON SHIPPING"
            elif row[1] == "WAYFAIR":
                row[20] = "WAYFAIR"
            elif row[1] == "VWR INTERNATIONAL":
                row[20] = "UPS GROUND"
            elif row[1] == "FISHER SCIENTIFIC":
                row[20] = "TRIPPNT GROUND"

            """Address check functionality begins here---------------------------------------------------------------"""
            if not (row[1] == "AMAZON" or row[4] == "VWR International"):
                allAddressLines = [row[4], row[5], row[6], row[41], row[54], row[56], row[58]]
                i = 0
                uniqueAddressLines = [[allAddressLines[0], 0]]
                for line in allAddressLines:
                    unique = True
                    for uniqueLine in uniqueAddressLines:
                        if uniqueLine[0] in line:
                            unique = False
                            uniqueLine[0] = line
                        elif line in uniqueLine[0]:
                            unique = False
                    if unique:
                        uniqueAddressLines.append([line, i])
                    i += 1
                addressLines = []
                uniqueIndexes = []
                for line in uniqueAddressLines:
                    if line[1] in [0, 1, 2, 4, 5]:
                        addressLines.append(line[0])
                attnLine = ""
                for line in uniqueAddressLines:
                    if line[1] == 3:
                        attnLine = attnLine + " REF: "+line[0]
                    if line[1] == 6:
                        if len(attnLine) >= 1:
                            attnLine = attnLine + "/" + line[0]
                        else:
                            attnLine = " ATTN: " + line[0]
                if len(addressLines) == 1:
                    row[4] = addressLines[0]
                    row[5] = ""
                    row[6] = ""
                    row[54] = ""
                    row[56] = ""
                elif len(addressLines) == 2:
                    row[4] = addressLines[0]
                    row[5] = addressLines[1]
                    row[6] = ""
                    row[54] = ""
                    row[56] = ""
                elif len(addressLines) == 3:
                    row[4] = addressLines[0]
                    row[5] = addressLines[1]
                    row[6] = addressLines[2]
                    row[54] = ""
                    row[56] = ""
                elif len(addressLines) == 4:
                    row[4] = addressLines[0]
                    row[5] = addressLines[1]
                    row[6] = addressLines[2]
                    row[54] = addressLines[3]
                    row[56] = ""
                elif len(addressLines) == 5:
                    row[4] = addressLines[0]
                    row[5] = addressLines[1]
                    row[6] = addressLines[2]
                    row[54] = addressLines[3]
                    row[56] = addressLines[4]
                if row[10] == "CA":
                    countryCode = "CA"
                else:
                    countryCode = "US"
                postalCode = row[9]
                addressesToValidate = []
                cleanedAddressLines = []
                # clean address lines and add them to the list to be validated
                for item in addressLines:
                    item = item.replace("-", "")
                    item = item.replace(",", " ")
                    item = item.replace(".", "")
                    cleanedAddressLines.append(item)
                    address = {'streetLines': [item, "this line left intentionally blank"], 'postalCode': postalCode,
                               'countryCode': countryCode}
                    addressesToValidate.append({"address": address})
                # ask FedEx to validate the lines
                response = validateAddress(addressesToValidate)
                # check the response to get the corrected street address
                a = 0
                streetAddress = row[5]
                streetAddress2 = []
                try:
                    resolvedAddresses = response.json()["output"]["resolvedAddresses"]
                except KeyError:
                    a = 99
                    resolvedAddresses = []
                while a < len(resolvedAddresses):
                    try:
                        classification = response.json()["output"]["resolvedAddresses"][a]["classification"]
                        if not classification == "UNKNOWN":
                            streetAddress = response.json()["output"]["resolvedAddresses"][a]["streetLinesToken"][0]
                            try:
                                streetAddress2 = response.json()["output"]["resolvedAddresses"][a]["streetLinesToken"][1]
                            except IndexError:
                                streetAddress2 = []
                            break
                        if a == len(resolvedAddresses) - 1:
                            a = 99
                    except:
                        a = 99
                        break
                    a += 1
                # if a resolved address token is found, do the following:
                if not a >= 99:
                    # removes the FIRST instance of the character found in the resolved address from the
                    # corresponding cleaned address line
                    lineToReplace = cleanedAddressLines[a].upper()
                    for char in streetAddress:
                        lineToReplace = lineToReplace.replace(char, "", 1)
                    if len(streetAddress2) > 0:
                        for char in streetAddress2:
                            lineToReplace = lineToReplace.replace(char, "", 1)
                    if len(lineToReplace.replace(" ", "")) > 0:
                        kartsNotes[PO_Number] = "Extra address text: "+lineToReplace
                    addressLines[a] = lineToReplace
                    # removes any blank lines
                    while "" in addressLines:
                        addressLines.remove("")
                    # removes any duplicate lines
                    finalAddressLines = []
                    for item in addressLines:
                        if item.upper() not in finalAddressLines:
                            finalAddressLines.append(item.upper())
                    # adds the resolved address line(s) at the correct index(es)
                    finalAddressLines.insert(1, streetAddress)
                    if len(streetAddress2) > 0:
                        finalAddressLines.insert(2, streetAddress2)
                    row[4] = addressLines[0]
                    if len(addressLines) > 1:
                        row[5] = addressLines[1]
                        if len(addressLines) > 2:
                            row[6] = addressLines[2]
                            if len(addressLines) > 4:
                                row[54] = addressLines[3] + "/" + addressLines[4]
                            elif len(addressLines) > 3:
                                row[54] = addressLines[3]
                    row[4] = finalAddressLines[0]
                    while len(finalAddressLines) <= 5:
                        finalAddressLines.append("")
                    row[5] = finalAddressLines[1]
                    row[6] = finalAddressLines[2]
                    if finalAddressLines[4] == "":
                        row[54] = finalAddressLines[3] + attnLine
                    else:
                        row[54] = finalAddressLines[3] + "/" + finalAddressLines[4]+attnLine
                else:
                    kartsNotes[PO_Number] = "Address not Found"
                    row[54] = row[54] + attnLine
            """Price checking begins here----------------------------------------------------------------------------"""
            # Looks up the line item MSRP in the price dictionary. If no price is found, assumes product is obsolete.
            try:
                if customer == "AMAZON":
                    List_Price = amzn_price_dict[partNum]
                    # Special Amazon pricing document -- contains Amazon pre-discounted prices.
                else:
                    List_Price = price_dict[partNum]
            except KeyError:
                # Behavior if the item is NOT in the pricing dictionary:
                if customer in send855List:
                    # corrects for parts like 5XXXXBLUE, but still needs the first five characters to be integers.
                    if not is_integer(partNum):
                        partNum = 00000
                    # if the part number isn't OEM, generates an 855
                    if int(partNum) < 80000:
                        if len(row[33]) > 5:
                            partNum = row[33][0:5]
                            row[33] = partNum
                        else:
                            partNum = row[33]
                        POALine = ["855", customer, "Original", "Acknowledge - With Detail and Change", PO_Number,
                                   po_date.strftime("%m/%d/%Y"), "", "", "", "TrippNT", "10991 N Airworld Dr", "",
                                   "Kansas City", "MO", "64153", "", "Airworld", "", row[32], row[33], "", "", row[37],
                                   "ea", row[39], "", "", "", "", shipDate, "", "Item Rejected", row[37], "Each",
                                   "Canceled due to discontinued Item", row[4], row[5], row[6], row[7], row[8],
                                   shipZip, shipCountry, ]
                        if customer == "AMAZON DIRECT":
                            POALine[3] = "Reject with Detail"
                        POAUpload.append(POALine)
                    # If part number is OEM, kicks out portal rejection document for human review.
                    else:
                        rejectLine = [customer, PO_Number, partNum, "Obsolete Part"]
                        portal_rejected.append(rejectLine)
                    continue
                # If the customer is set to "portal reject," kicks out portal rejection document for human review.
                if customer in portalRejectList:
                    rejectLine = [customer, PO_Number, partNum, "Obsolete Part"]
                    portal_rejected.append(rejectLine)
                    continue
                else:
                    # Sends email to customer about a discontinued product.
                    rejected.append(PO_Number)
                    reject_message = generate_message(row[12], PO_Number, "Discontinued", partNum, PO_Price)
                    rejectReason.append(reject_message)
                    reject_emails.append(get_email(row[1]), customerAddresses)
                    subject_reasons.append("Obsolete Product")
                continue
            # Calculates the customer's discounted price. If discount is not defined, assumes the PO price is correct.
            if customer == "FISHER SCIENTIFIC" or customer == "WAYFAIR" or customer == "AMAZON DIRECT":
                Dist_Price = round(List_Price * (1 - .275), 2)
            elif customer == "GLOBAL":
                if int(partNum[0]) == 9:
                    Dist_Price = List_Price
                else:
                    Dist_Price = round(List_Price * 0.65, 2)
            elif customer == "VWR INTERNATIONAL":
                if int(partNum[0]) == 9:
                    Dist_Price = List_Price
                else:
                    Dist_Price = round(List_Price / 1.65, 2)-0.26
                    # subtracting 0.25 here effectively lets VWR prices be up to 31 cents low. Adjusts for rounding.
            elif customer == "AMAZON":
                Dist_Price = List_Price
                deliverDate = row[27]
            else:
                Dist_Price = PO_Price
            if customer in fixPrice:
                row[39] = Dist_Price
                PO_Price = float(row[39])
            # Rejects any prices that are not within the bounds (-$0.05 to + 10%)
            if not Dist_Price - .05 <= PO_Price and not webOrder:
                if customer in send855List:
                    if int(partNum) < 90000:
                        POALine = ["855", customer, "Original", "Acknowledge - With Detail and Change", PO_Number,
                                   po_date.strftime("%m/%d/%Y"), "", "", "", "TrippNT", "10991 N Airworld Dr", "",
                                   "Kansas City", "MO", "64153", "", "Airworld", "", row[32], row[33], "", "", row[37],
                                   "ea", row[39], "", "", "", "", shipDate, "", "Item Rejected", row[37], "Each",
                                   "Canceled due to missing/Invalid Unit Price", row[4], row[5], row[6], row[7], row[8],
                                   shipState, shipZip, shipCountry]
                        if customer == "AMAZON DIRECT":
                            POALine[3] = "Reject with Detail"
                        POAUpload.append(POALine)
                    else:
                        rejectLine = [customer, PO_Number, partNum, "Obsolete Part"]
                        portal_rejected.append(rejectLine)
                    continue
                if customer in portalRejectList:
                    rejectLine = [customer, PO_Number, partNum, "Price Low"]
                    portal_rejected.append(rejectLine)
                    continue
                else:
                    # send an email to customer about incorrect pricing
                    rejected.append(PO_Number)
                    reject_message = generate_message(row[12], PO_Number, "Pricing", partNum, Dist_Price)
                    rejectReason.append(reject_message)
                    reject_emails.append(get_email(row[1]), customerAddresses)
                    subject_reasons.append("Pricing Discrepancy")
            # if the PO_Price is too high (more than 10%)
            elif Dist_Price * 1.1 < PO_Price and not webOrder:
                # Sends portal rejection for these specific customers
                if customer in portalRejectList:
                    rejectLine = [customer, PO_Number, partNum, "Price High"]
                    portal_rejected.append(rejectLine)
                    continue
                else:
                    # Generates the rejection email for non specified Customers
                    rejected.append(PO_Number)
                    reject_message = generate_message(row[12], PO_Number, "Pricing", partNum, Dist_Price)
                    rejectReason.append(reject_message)
                    reject_emails.append(get_email(row[1]), customerAddresses)
                    subject_reasons.append("Pricing Discrepancy")
            processing.append(row[x])
            # Handles customers in 855 list and appends to the list
            if customer in send855List:
                POALine = ["855", customer, "Original", "Acknowledge - With Detail: No Change", PO_Number,
                           po_date.strftime("%m/%d/%Y"), "", "", "", "TrippNT", "10991 N Airworld Dr", "",
                           "Kansas City", "MO", "64153", "", "Airworld", "", row[32], row[33], "", "", row[37], "ea",
                           row[39],
                           "", "", deliverDate, "", shipDate, "", "Item Accepted", row[37], "Each",
                           "Shipping 100 percent of ordered product", row[4], row[5], row[6], row[7], row[8],
                           shipZip, shipCountry]
                if customer == "AMAZON DIRECT":
                    POALine[3] = "Accepted"
                POAUpload.append(POALine)
        else:
            # creates a dictionary of the header so that columns can be called by name instead of index
            columnCounter = 0
            for columnName in row:
                if columnName in tcIndex.keys():
                    columnName = columnName.lower()
                tcIndex[columnName] = columnCounter
                columnCounter += 1
            processing.append(row[x])
        counter += 1
counter = 0
line_total = 0
order_total = 0
orderWeight = 0
orderCube = 0
Processing2 = []

"""Total-dependent items are calculated here ------------------------------------------------------------------------"""
# iterates over processing and calculates the order total price, weight, and volume
ltlOverride = {}
orderCount = 0
while counter < len(processing):
    if counter > 0:
        current_line = processing[counter]
        previous_line = processing[counter - 1]
        customer = previous_line[1]
        line_total = int(current_line[37]) * float(current_line[39])
        ltlOverride[current_line[2]] = False
        try:
            partDims = weightsNdimsDict[current_line[33]]
        except KeyError:
            partDims = {"width": 1, "depth": 1, "height": 1, "weight": 1, "ltlThreshold": ""}
            if current_line[2] in list(kartsNotes.keys()):
                kartsNotes[current_line[2]] += "Can't find dims for part " + str(current_line[33]) + "."
            else:
                kartsNotes[current_line[2]] = "Can't find dims for part " + str(current_line[33]) + "."
        if not partDims["ltlThreshold"] == "":
            if int(partDims["ltlThreshold"]) <= int(current_line[37]):
                ltlOverride[current_line[2]] = True
        lineWeight = float(partDims["weight"]) * int(current_line[37])
        linePacking = howManyFit([partDims["width"], partDims["depth"], partDims["height"]], int(current_line[37]),
                                 boxes, pad)
        if linePacking["qtyPerPack"] == 0:
            lineCube = (partDims["width"] * partDims["depth"] * partDims["height"])/(12*12*12)
        else:
            lineCube = (math.ceil(int(current_line[37])/linePacking["qtyPerPack"]) * partDims["width"]
                        * partDims["depth"] * partDims["height"]) / (12 * 12 * 12)
        # adds previous line to current line if they are the same PO
        if current_line[2] == previous_line[2]:
            order_total += line_total
            orderWeight += lineWeight
            orderCube += lineCube
        else:
            # Calculates the shipping fee
            if order_total > 0:
                freight = False
                if orderDims["weight"] > 150 or orderDims["volume"] > 60 or ltlOverride[previous_line[2]]:
                    freight = True
                orderCount += 1
                createSQL(previous_line, freight, orderCount)
                shipping_cost = 0
                shipping_line = []
                shipping_line += previous_line
                if customer == "VWR INTERNATIONAL":
                    shipping_cost = round(order_total * .02, 2)
                    shipping_line[33] = "XDROP"
                    if freight:
                        for line in processing:
                            if line[2] == previous_line[2]:
                                line[20] = "LTL"
                        shipping_line[20] = "LTL"
                elif customer == "FISHER SCIENTIFIC":
                    shipping_cost = round(order_total * .3, 2) + 20
                    if freight:
                        shipping_cost += 80
                        for line in processing:
                            if line[2] == previous_line[2]:
                                line[20] = "TRIPPNT FREIGHT"
                        shipping_line[20] = "TRIPPNT FREIGHT"
                    shipping_line[33] = "SHIPPING"
                elif customer == "GLOBAL":
                    if not len(shipping_line[20].replace("GROUND", "")) < len(shipping_line[20]) \
                            or len(shipping_line[20].replace("NFI", "")) < len(shipping_line[20]):
                        if shipping_line[2] in list(kartsNotes.keys()):
                            kartsNotes[shipping_line[2]] += "Ship via " + shipping_line[20]
                        else:
                            kartsNotes[shipping_line[2]] = "Ship via " + shipping_line[20]
                shipping_line[32] = int(shipping_line[32]) + 1
                shipping_line[34] = shipping_line[33]
                shipping_line[35] = ""
                shipping_line[36] = ""
                shipping_line[38] = "Each"
                shipping_line[40] = ""
                shipping_line[37] = 1
                shipping_line[39] = shipping_cost
                Processing2.append(shipping_line)
            order_total = line_total
            orderWeight = lineWeight
            # orderCube = lineCube
        total_dict[current_line[2]] = order_total
        orderDims = {"weight": orderWeight, "volume": orderCube}
        orderDimsDict[current_line[2]] = orderDims
        # finds the last line item to end the loop
        if counter == (len(processing) - 1):
            freight = False
            if orderDims["weight"] > 150 or orderDims["volume"] > 60 or ltlOverride[current_line[2]]:
                freight = True
            orderCount += 1
            createSQL(current_line, freight, orderCount)
            customer = current_line[1]
            shipping_line = []
            shipping_line += current_line
            shipping_cost = 0
            if customer == "VWR INTERNATIONAL":
                shipping_cost = round(order_total * .02, 2)
                shipping_line[33] = "XDROP"
                if freight:
                    for line in processing:
                        if line[2] == current_line[2]:
                            line[20] = "LTL"
                    shipping_line[20] = "LTL"
            elif customer == "FISHER SCIENTIFIC":
                shipping_cost = round(order_total * .3, 2) + 20
                if freight:
                    shipping_cost += 80
                    for line in processing:
                        if line[2] == current_line[2]:
                            line[20] = "TRIPPNT FREIGHT"
                    shipping_line[20] = "TRIPPNT FREIGHT"
                shipping_line[33] = "SHIPPING"
            elif customer == "GLOBAL":
                if not len(shipping_line[20].replace("GROUND", "")) < len(shipping_line[20]) \
                        or len(shipping_line[20].replace("NFI", "")) < len(shipping_line[20]):
                    if shipping_line[2] in list(kartsNotes.keys()):
                        kartsNotes[shipping_line[2]] += "Ship via "+shipping_line[20]
                    else:
                        kartsNotes[shipping_line[2]] = "Ship via "+shipping_line[20]
            shipping_line[32] = int(shipping_line[32]) + 1
            shipping_line[34] = shipping_line[33]
            shipping_line[35] = ""
            shipping_line[36] = ""
            shipping_line[38] = "Each"
            shipping_line[40] = ""
            shipping_line[37] = 1
            shipping_line[39] = shipping_cost
            Processing2.append(shipping_line)
    counter += 1
# just resetting variables
counter = 0
errors = 0
"""Output (CSV and email) begins-------------------------------------------------------------------------------------"""
posWithAlerts = list(kartsNotes.keys())[:]
notesCounter = 0
counter2 = 0
for item in processing:
    if counter2 > 0 and not webOrder:
        if rejected.count(item[2]) == 0:
            # Try loop is checking for unicode errors.
            # I would check in read, but I can't move the file to error while it's still open.
            try:
                # removes buyer info from Fisher notes section (per customer request)
                if item[1] == "FISHER SCIENTIFIC":
                    item[23] = ""
                    i = 42
                    while i <= 45:
                        item[i] = ""
                        i += 1
                    # looks for any special instructions and writes them to a "notes" file for the dashboard
                    if not item[40] == "":
                        if notesCounter == 0:
                            notesOutput = open(
                                "H:\\Order Entry\\EDI 2\\Notes" + datetime.fromtimestamp(file_creation_date).strftime(
                                    "%d_%H_%M") +
                                ".csv", 'w', newline='')
                            notesWriter = csv.writer(notesOutput)
                        notes = [item[2], item[40]]
                        if item[2] in list(kartsNotes.keys()):
                            notes[1] = notes[1]+", "+kartsNotes[item[2]]
                            del kartsNotes[item[2]]
                        notesWriter.writerow(notes)
                        notesCounter += 1
                # looks for any special instructions and writes them to a "notes" file for the dashboard
                elif item[1] == "VWR INTERNATIONAL":
                    if not item[55] == "":
                        if notesCounter == 0:
                            notesOutput = open(
                                "H:\\Order Entry\\EDI 2\\Notes" + datetime.fromtimestamp(file_creation_date).strftime(
                                    "%d_%H_%M") + ".csv", 'w', newline='')
                            notesWriter = csv.writer(notesOutput)
                        notes = [item[2], item[55]+": "+item[40]]
                        if item[2] in list(kartsNotes.keys()):
                            notes[1] = notes[1]+", "+kartsNotes[item[2]]
                            del kartsNotes[item[2]]
                        notesWriter.writerow(notes)
                        notesCounter += 1
                if len(kartsNotes) > 0:
                    if notesCounter == 0:
                        notesOutput = open(
                            "H:\\Order Entry\\EDI 2\\Notes" + datetime.fromtimestamp(file_creation_date).strftime(
                                "%d_%H_%M") + ".csv", 'w', newline='')
                        notesWriter = csv.writer(notesOutput)
                        notesCounter += 1
                    for key in list(kartsNotes.keys()):
                        notes = [key, kartsNotes[key]]
                        notesWriter.writerow(notes)
                if item[33] == 00000:
                    item[33] = "SHIPPING"
                e2item = list(convertToE2(item, tcIndex, webOrder, posWithAlerts))
                writer.writerow(e2item)
            except UnicodeEncodeError:
                shutil.move(maxFile, "C:\\True Commerce\\Transaction Manager\\Export\\Error")
                output = open("H:\\Order Entry\\EDI 2\\ERROR " +
                              datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M") + ".csv", 'w', newline='')
                writer = csv.writer(output)
                writer.writerow("Unicode error detected at " + todays_date.strftime("%H:%M"))
                output.close()
                raise SystemExit
    # This just writes the header row. It can't go through the main "if" because the data format is not expected.
    elif webOrder and counter2 > 0:
        item[33] = item[34]
        e2item = list(convertToE2(item, tcIndex, webOrder, posWithAlerts))
        writer.writerow(e2item)
    else:
        headerRow = list(e2Index.keys())
        headerRow.insert(45, "SHIP TO NAME")
        writer.writerow(headerRow)
    counter2 += 1
# closes notes, if it was opened.
if notesCounter > 0:
    notesOutput.close()
r = 0
counter = 0
# writes the drop ship fee for VWR drop ship orders
while counter < len(Processing2):
    if Processing2[counter][1] == "VWR INTERNATIONAL":
        if not ("VWR" in Processing2[counter][4]):
            e2item = list(convertToE2(Processing2[counter], tcIndex, webOrder, posWithAlerts))
            writer.writerow(e2item)
    elif Processing2[counter][1] == "FISHER SCIENTIFIC":
        e2item = list(convertToE2(Processing2[counter], tcIndex, webOrder, posWithAlerts))
        writer.writerow(e2item)
    counter += 1

# Generates reject emails, if any rejected PO's match the defined rejection email format(s).
# Writes any remaining rejections to "rejected" file for human attention.
if len(rejected) > 0:
    with open("H:\Rejection Emails\Rejected" + datetime.fromtimestamp(file_creation_date).strftime("%m_%d_%H_%M")
              + ".csv", 'w',
              newline='') as Reject_Data:
        dialect = csv.excel()
        writer2 = csv.writer(Reject_Data, dialect)
        rejectLine = []
        start_line = "Email", "Subject Line", "Message"
        writer2.writerow(start_line)
        # sends rejection email if only 1 - 2 errors but puts it in rejection folder for 3+ errors
        while r < len(rejected):
            subject_line = "PO " + rejected[r] + " has been Canceled: " + subject_reasons[r]
            linesRejected = rejected.count(rejected[r])
            if testMode:
                reject_emails[r] = "<Sales@TrippNT.com>"
            if subject_reasons[r] == "Pricing Discrepancy":
                # generates the normal message
                if linesRejected == 1:
                    message = MIMEmail("TrippNT DoNotReply <KARTS@trippnt.com>", reject_emails[r], subject_line,
                                       rejectReason[r])
                # generates double rejection message
                # index messageBody backwards to form an iterative for variable lines?
                elif linesRejected == 2 and subject_reasons[r + 1] == "Pricing Discrepancy":
                    messageBody = rejectReason[r][0:rejectReason[r].find("USD") + 4] + \
                                  rejectReason[r + 1][rejectReason[r + 1].find("production.") + 13:]
                    message = MIMEmail("TrippNT DoNotReply <KARTS@trippnt.com>", reject_emails[r], subject_line,
                                       messageBody)
                    r += 1
                else:
                    rejectLine.append(reject_emails[r])
                    rejectLine.append(subject_line)
                    rejectLine.append(rejectReason[r])
                    writer2.writerow(rejectLine)
                    rejectLine = []
                    # skip the email-sending bit if conditions aren't met
                    continue
                encoded_email = message.as_string()
                # encodes and trys to send the email if failed will send to rejection folder
                try:
                    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
                        server.login(sender_email, password)
                        server.sendmail(message["From"], message["To"], encoded_email)
                        # Code was being broken here "message["Bcc"],"
                        server.sendmail(message["From"], message["Bcc"], encoded_email)
                    reject_emails.append(get_email(row[1]), customerAddresses)
                    subject_reasons.append("Pricing Discrepancy")
                except:
                    rejectLine.append(reject_emails[r])
                    rejectLine.append(subject_line)
                    rejectLine.append(rejectReason[r])
                    writer2.writerow(rejectLine)
                    rejectLine = []
            # rejection folder functionality
            else:
                rejectLine.append(reject_emails[r])
                rejectLine.append(subject_line)
                rejectLine.append(rejectReason[r])
                writer2.writerow(rejectLine)
                rejectLine = []
            r += 1
    Reject_Data.close()
# writes any portal rejections to a portal rejected file
r = 0
if len(portal_rejected) > 0:
    with open("H:\\Order Entry\\Portal Rejections\\Portal_Rejected" +
              datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M") + ".csv", 'w',
              newline='') as Portal_Reject_Data:
        dialect = csv.excel()
        writer3 = csv.writer(Portal_Reject_Data, dialect)
        start_line = "Customer", "PO", "Part Number", "Reason"
        writer3.writerow(start_line)
        while r < len(portal_rejected):
            writer3.writerow(portal_rejected[r])
            r += 1
    Portal_Reject_Data.close()
# closes the new PO and moves the TrueCommerce file to "Processed"
PO.close()
shutil.move(maxFile, "C:\\True Commerce\\Transaction Manager\\Export\\Processed")
# generates the 855 flat file
counter = 1
while len(POAUpload) > 0:
    if testMode:
        eight55 = open("H:\\Order Entry\\KARTS files\\Testing\\855Upload" +
                       datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M_") + str(counter) + ".csv",
                       'w', newline='')
    else:
        eight55 = open("C:\\True Commerce\\Transaction Manager\\Import\\855Upload" +
                       datetime.fromtimestamp(file_creation_date).strftime("%d_%H_%M_") + str(counter) + ".csv",
                       'w', newline='')
    writer = csv.writer(eight55)
    POAHeader = ["Transaction ID", "Accounting ID", "Purpose", "Type Status", "PO #", "PO Date", "Release Number",
                 "Request Ref Number", "Contract Number", "Selling Party Name", "Selling Party Address 1",
                 "Selling Party Address 2", "Selling Party City", "Selling Party State", "Selling Party Zip",
                 "Vendor Number", "Warehouse ID", "Line #", "PO Line #", "Vendor Part #", "UPC", "SKU", "Qty",
                 "UOM", "Price", "Scheduled Delivery Date", "Scheduled Delivery Time", "Estimated Delivery Date",
                 "Estimated Delivery Time", "Promised Date", "Promised Time", "Status", "Status Qty", "Status UOM",
                 "Cancel Reason", "Ship to Name", "Ship to Address 1", "Ship to Address 2", "Ship to City",
                 "Ship to State",
                 "Ship to Zip", "Ship to Country"]
    writer.writerow(POAHeader)
    currentPO = POAUpload[0][4]
    # deletes each line as it writes it to the new file
    for item in POAUpload:
        if item[4] == currentPO:
            writer.writerow(item)
            del POAUpload[0]
        else:
            break
    output.close()
    counter += 1
raise SystemExit
