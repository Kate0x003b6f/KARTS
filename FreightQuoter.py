import csv
import time
import math
import requests
import json

# item - type list - at least three indexes containing length, width, and height (order does not matter) of item
# qty - type int - qty of items to pack
# packageOptions - list of lists - one or more packages, in the same format as item, to choose between
# pad - type int - amount of space needed between the items and the edge of the package
def howManyFit(item, qty, packageOptions, pad):
    pad = pad*2
    selectedPackage = []
    qty1 = qty
    maxPerLayer = 0
    layerHeight = 0
    #Checks 6 different stacking options using l*w*H
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
                if maxPerPack == dir1 or maxPerPack == dir5:
                    layerHeight = item[2]
                    maxPerLayer = maxPerPack/math.floor((package[2] - pad) / item[2])
                elif maxPerPack == dir2 or maxPerPack == dir6:
                    layerHeight = item[0]
                    maxPerLayer = maxPerPack / math.floor((package[2] - pad) / item[0])
                else:
                    layerHeight = item[1]
                    maxPerLayer = maxPerPack / math.floor((package[2] - pad) / item[1])
                break
        if len(selectedPackage) == 0:
            qty1 -= 1
    # Calculates qty in each box
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
              # "lastPackQty": lastPackQty,
              "lastPackQty": 0,
              "layerHeight": layerHeight,
              "maxPerLayer": maxPerLayer
              }
    return output

pad = 2 # amount of padding, in inches, needed on each side of the box.
palletPad = 0.5 # amount of padding, in inches, needed on each side of the pallet.
maxGroundWeight = 150
maxGroundCube = 60
lineItemCount = 1
lineItems = []
with open("C:\\Scripts\\Stored Data\\Weights_n_Dims.csv", 'r') as weightsNdims:
    weightsNdimsDict = {}
    rcsv = csv.reader(weightsNdims, delimiter=',')
    counter = 0
    for row in rcsv:
        if counter > 0:
            weightsNdimsDict[row[0]] = {"width": float(row[1]), "height": float(row[2]), "depth": float(row[3]),
                                        "weight": float(row[4])}
        counter += 1
    weightsNdims.close()
# our box sizes (may need to update if we change box sizes)
# Doesn't know about telescope boxes
boxes = [[12, 12, 6], [12, 12, 12], [14, 14, 14], [20, 14, 12], [26, 14, 16], [18, 18, 18], [27, 13, 23], [21, 21, 21],
         [24, 24, 24]]
# our most common pallet sizes, with an arbitrary height of 66 inches to satisfy FedEx Express Freight
pallets = [[30, 50, 66, 30], [40, 48, 66, 35]]
# Variables to control looping behavior for user input
validAnswer = 0
shipmentComplete = False
shipmentTotal = 0
while not shipmentComplete:
    validAnswer = 0
    while validAnswer == 0:
        partNum = input("Please enter the five digit TrippNT part number for the item you want to ship: ")
        try:
            partDims = weightsNdimsDict[partNum]
            validAnswer = 1
        except KeyError:
            print("I can't seem to find that...")
            time.sleep(1)
    print("Here is what I have for part "+partNum+": ", partDims)
    validAnswer = 0
    while validAnswer == 0:
        # the dims currently contain a mix of packaging dims (for items that always ship one to a box) and item dims.
        boxed = input("Is this item already in a box? Please type y for yes or n for no: ")
        boxed = boxed.lower()
        if boxed == "y" or boxed == "n":
            validAnswer = 1
        else:
            print("Sorry, I didn't understand that.")
            time.sleep(1)
    validAnswer = 0
    while validAnswer == 0:
        qty = input("Thank you. Now, how many "+partNum+" are on this shipment? Please enter an integer (ex 1): ")
        try:
            qty = int(qty)
            validAnswer = 1
        except ValueError:
            print("Hm, that doesn't look like an integer. Please try again. ")

    if boxed == "n":
        shipment = howManyFit([partDims["width"], partDims["height"], partDims["depth"]], qty, boxes, pad)
        qtyPerBox = shipment["qtyPerPack"]
        selectedBox = shipment["selectedPackage"]
        lastBoxQty = shipment["lastPackQty"]
        if qtyPerBox > 0:
            print(".\n.")
            print(qtyPerBox, "ea", partNum, "should fit in a", selectedBox, "box with", pad, "inch(es) of padding per side.")
            if qtyPerBox < qty:
                qtyBoxes = math.ceil(qty/qtyPerBox)
                print("You will need", qtyBoxes, "boxes to fit all", qty, partNum)
                lastBoxWeight = partDims["weight"] * lastBoxQty + selectedBox[0]/12
            else:
                qtyBoxes = 1
            boxWeight = partDims["weight"] * qtyPerBox + selectedBox[0]/12
            print(qtyPerBox, partNum, ", plus a", selectedBox, "box, should weigh about", boxWeight, "pounds.")
            if not lastBoxQty == 0:
                print("The final box will contain", lastBoxQty, partNum)
                print("It will weigh around", lastBoxWeight, "pounds")
                print("\nNOTE -- while I know how to calculate this, when I SHIP, I will assume all boxes are full.\n")
        else:
            print("\nHm. I don't think this item will fit in a standard box with", pad, "inches of padding.\n")
            time.sleep(1)
            print("My programming doesn't appear to cover this scenario. Time to panic! Mission abort!! ABORRT!!!"
                  "\n...\nGoodbye! See you again soon! :)")
            raise SystemExit
    else:
        qtyBoxes = qty
        selectedBox = [partDims["width"], partDims["depth"], partDims["height"]]
        boxWeight = partDims["weight"]
        lastBoxQty = 0
    validAnswer = 0
    lineItems.append([qtyBoxes, selectedBox, boxWeight])
    while validAnswer == 0:
        addAnother = input("\nAdd another line item? Please type y for yes or n for no: ")
        if addAnother.lower() == "y":
            lineItemCount += 1
            validAnswer = 1
        elif addAnother.lower() == "n":
            shipmentComplete = True
            validAnswer = 1
        else:
            print("Sorry, I didn't understand that. Please try again.")
shipmentWeight = 0
shipmentDimWeight = 0
# calculates the weight and dimweight for each line item, then totals them
for item in lineItems:
    lineWeight = item[0]*item[2]
    lineDimWeight = (item[1][0]/12)*(item[1][1]/12)*(item[1][2]/12)/139
    shipmentWeight += lineWeight
    shipmentDimWeight += lineDimWeight
# determines whether or not to palletize the shipment. (It's not very good at palletizing shipments, FYI)
if shipmentWeight > maxGroundWeight or shipmentDimWeight > maxGroundCube:
    print("This should probably ship on a pallet.")
    if lineItemCount > 1:
        print("I don't know how to palletize different boxes together. Goodbye.")
        raise SystemExit
    ground = False
    # Pallet fitting
    shipment = howManyFit(selectedBox, qtyBoxes, pallets, palletPad)
    qtyPerPallet = shipment["qtyPerPack"]
    selectedPallet = shipment["selectedPackage"]
    lastPalletQty = shipment["lastPackQty"]
    layerHeight = shipment["layerHeight"]
    maxPerLayer = shipment["maxPerLayer"]
    layersNeeded = math.ceil(qtyPerPallet/maxPerLayer)
    palletHeight = layersNeeded * layerHeight + 6
    selectedPallet[2] = palletHeight
    if qtyPerPallet > 0:
        print(".\n.")
        print(qtyPerPallet, "ea", selectedBox, "boxes should fit on a", selectedPallet[0:2], "pallet with", palletPad,
              "inch(es) of padding per side. It will be", palletHeight, "inches tall.")
        if layersNeeded > 1:
            print(layersNeeded, "layers will be on the pallet. Each layer will be", layerHeight, "inches high.")
        if qtyPerPallet <= qtyBoxes:
            qtyPallets = math.ceil(qtyBoxes/qtyPerPallet)
            print("You will need", qtyPallets, "pallets to fit all", qtyBoxes, selectedBox, "boxes.")
            lastPalletQty = qtyBoxes % qtyPerPallet
            lastPalletWeight = boxWeight * lastPalletQty + 35
            lastPalletHeight = layerHeight * math.ceil(lastPalletQty/maxPerLayer)
            if palletHeight + layerHeight < 90:
                print("\n*** This shipment could most likely be one layer higher, but that would require"
                      " pre-approval from FedEx, as it would exceed 70 inches in height.***\n")
        else:
            qtyPallets = 1
            lastPalletWeight = 0
        palletWeight = qtyPerPallet * boxWeight + 35
        print(qtyPerPallet, selectedBox, "boxes, plus a", selectedPallet[0:2], "pallet, should weigh about", palletWeight,
              "pounds.")
        if not lastPalletQty == 0:
            print("The final pallet will have", lastPalletQty, selectedBox, "boxes.")
            print("It will weigh around", lastPalletWeight, "pounds")
            print("\nNOTE -- while I know how to calculate this, when I SHIP, I will assume all pallets are full.\n")
    else:
        print("\nHm. I don't think this item will fit on a standard pallet with", palletPad, "inches of padding.\n")
        time.sleep(1)
        print("My programming doesn't appear to cover this scenario. Time to panic! Mission abort!! ABORRT!!!"
              "\n...\nGoodbye! See you again soon! :)")
        raise SystemExit
else:
    ground = True

print("Contacting FedEx...")
# get FedEx API token using credentials
url = "https://apis.fedex.com/oauth/token"

payload = {'grant_type': 'client_credentials', 'client_id': 'l7318f8a1517064b149055f02eafc65b70',
           'client_secret': 'f1d85ae4-2b7a-4642-b6b1-313250fbbf79'}
headers = {
    'Content-Type': "application/x-www-form-urlencoded"
    }
response = requests.request("POST", url, data=payload, headers=headers)
# wait for response to finish
time.sleep(1)
# extract token from response
token = "Bearer "+response.json()["access_token"]

# more input loops
validAnswer = 0
while validAnswer == 0:
    shipZip = input("Please enter the five-digit zip code this shipment is headed to: ")
    try:
        placeholder = int(shipZip)
        if len(shipZip) == 5:
            validAnswer = 1
        else:
            print("Please enter exactly five digits.")
    except ValueError:
        print("That doesn't look like a number. Please try again.")
print("\n")
i = 0
# generate JSON data for each line in the format FedEx is expecting
while i < lineItemCount:
    if ground:
        accountNum = 289152784
        packType = "BOX"
        carrierCode = "FDXG"
        requestedPackageLineItems = []
        lineItem = lineItems[i]
        info = {
            "subPackagingType": packType,
            "groupPackageCount": lineItem[0],
            "weight": {
                "units": "LB",
                "value": lineItem[2]
            },
            "dimensions": {
                "length": lineItem[1][0],
                "width": lineItem[1][1],
                "height": lineItem[1][2],
                "units": "IN"
            }
        }
        requestedPackageLineItems.append(info)

    else:
        accountNum = 675731896
        qtyPacks = qtyPallets
        packWeight = palletWeight
        selectedPack = selectedPallet
        packType = "PALLET"
        carrierCode = "FXFR"
        requestedPackageLineItems = [
          {
            "subPackagingType": packType,
            "groupPackageCount": qtyPacks,
            "weight": {
              "units": "LB",
              "value": packWeight
            },
            "dimensions": {
                "length": selectedPack[0],
                "width": selectedPack[1],
                "height": selectedPack[2],
                "units": "IN"
            }
          }
        ]

    payload = {
      "accountNumber": {
        "value": accountNum
      },
      "requestedShipment": {
        "shipper": {
          "address": {
            "postalCode": 64153,
            "countryCode": "US"
          }
        },
        "recipient": {
          "address": {
            "postalCode": shipZip,
            "countryCode": "US"
          }
        },
        "pickupType": "USE_SCHEDULED_PICKUP",
        "rateRequestType": [
          "ACCOUNT"
        ],
        "requestedPackageLineItems": requestedPackageLineItems,
        "serviceTypeDetail": {
            "carrierCode": carrierCode
        },
        "groundShipment": ground
      }
    }
    # encode and send request (one per line item)
    payload = json.dumps(payload).encode('utf8')
    url = "https://apis.fedex.com/rate/v1/rates/quotes"

    headers = {
        'Content-Type': "application/json",
        'X-locale': "en_US",
        'Authorization': token
        }

    response = requests.request("POST", url, data=payload, headers=headers)
    quotes = []
    quotesCount = 0
    services = {}
    # process the response to extract the actual quoted amount
    while quotesCount < len(response.json()["output"]["rateReplyDetails"]):
        quotes.append(response.json()["output"]["rateReplyDetails"][quotesCount]
                      ["ratedShipmentDetails"][0]["totalNetFedExCharge"])
        services[quotes[-1]] = response.json()["output"]["rateReplyDetails"][quotesCount]["serviceType"]
        quotesCount += 1
    # find the cheapest service for this shipment
    selectedService = services[min(quotes)]
    shipmentTotal += min(quotes)
    print("Price to ship package", i+1, "via", selectedService, ": ", "$%.2f" % min(quotes))
    i += 1
print("Price for all packages with 20% handling charge: ", "$%.2f" % (shipmentTotal * 1.2))
if not ground:
    print("\n*** Cheaper freight services, such as FedEx Freight Economy, may be available through the toolbox. ***")
print("\nSee you next time!")
