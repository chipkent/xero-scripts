#!/usr/bin/env python
# 
# Converts an Amazon order history download into a CSV to import into Xero
# Generate Amazon report (Items): https://www.amazon.com/gp/b2b/reports
# Import into Xero:  https://go.xero.com/Accounts/Payable/Dashboard/
#

import sys
import csv
import logging

logging.basicConfig(level=logging.INFO)

def parse_key(row):
	return (row["Order ID"], row["Carrier Name & Tracking Number"])

def skip_row(row):
	return row["Order Status"] == "Cancelled" and row["Quantity"] == "0"

class OrderReport:
	def __init__(self, fileName):
		logging.info("Creating OrderReport(%s)",(fileName))
		self._map = {}
		with open(fileName) as csvfile:
			data = csv.DictReader(csvfile, delimiter=',', quotechar='"')
			for row in data:
				key = parse_key(row)
				value = {}

				for field in ["Subtotal", "Shipping Charge", "Tax Before Promotions", "Total Promotions", "Tax Charged", "Total Charged"]:
					value[field] = (row[field].replace('$',''))

				if key in self._map:
					raise Exception("Duplicate: ", key, self._map[key], value)

				self._map[key] = value

	def __getitem__(self, key):
		return self._map[key]



if len(sys.argv) !=3:
	print "Usage <amzn_item_report.csv> <amzn_order_report.csv>"
	sys.exit(-1)

fileNameItem = sys.argv[1]
fileNameOrder = sys.argv[2]
fileNameOutput = fileNameItem.replace('.csv','.NEW_TEST.csv')
logging.warn("fileNameItem: %s", fileNameItem)
logging.warn("fileNameOrder: %s", fileNameOrder)
logging.warn("fileNameOutput: %s", fileNameOutput)

orderReport = OrderReport(fileNameOrder)

reader = csv.DictReader(open(fileNameItem), delimiter=',', quotechar='"')
outfile = open(fileNameOutput,'wb') #TODO rename
writer = csv.writer(outfile, delimiter=',', quotechar='"')

newHeader = [
	"*ContactName",
	"EmailAddress",
	"POAddressLine1",
	"POAddressLine2",
	"POAddressLine3",
	"POAddressLine4",
	"POCity",
	"PORegion",
	"POPostalCode",
	"POCountry",
	"*InvoiceNumber",
	"*InvoiceDate",
	"*DueDate",
	"InventoryItemCode",
	"Description",
	"*Quantity",
	"*UnitAmount",
	"*AccountCode",
	"*TaxType",
	"TrackingName1",
	"TrackingOption1",
	"TrackingName2",
	"TrackingOption2",
	"Currency"	
	]

writer.writerow( newHeader )

#TODO restructure this section (OrderAdjustments?)
#TODO switch to logging for debug crap

orderBreakdown = {}

for row in reader:
	if skip_row(row):
		logging.debug("**** Skipping Row ****")
		logging.debug(row)
		continue

	key = parse_key(row)

	if key not in orderBreakdown:
		value = {"Quantity":0, "Item Total": 0.0, "Shipping soon": 0, "Item Subtotal Tax": 0.0}
		orderBreakdown[key] = value

	value = orderBreakdown[key]
	value["Quantity"] += float(row["Quantity"])
	value["Item Total"] += float(row["Item Total"].replace('$',''))
	value["Item Subtotal Tax"] += float(row["Item Subtotal Tax"].replace('$',''))
	
	if row["Order Status"] == "Shipping soon":
		value["Shipping soon"] += 1

	

for key in orderBreakdown:
	value = orderBreakdown[key]
	logging.debug("***************************************************")
	logging.debug("%s %s", key, value)

	try:
		#TODO Clean up this section
		orv = orderReport[key]
		#print orv
		taxAdjustment = float(orv["Tax Charged"]) - value["Item Subtotal Tax"] #TODO ???
		totalAdjustment = float(orv["Shipping Charge"]) - float(orv["Total Promotions"]) + taxAdjustment
		totalAdjustmentPerUnit = totalAdjustment / value["Quantity"]

		value["Tax Adjustment"] = taxAdjustment
		value["Total Adjustment"] = totalAdjustment
		value["Total Adjustment Per Unit"] = totalAdjustmentPerUnit

		v1 = float(value["Item Total"]) + totalAdjustment
		v2 = float(orv["Total Charged"])
		logging.debug("Consistency Check: %s %s %s %s %s %s %s %s", key, v1, v2, abs(v1-v2)<1e-5, value["Item Subtotal Tax"], orv["Tax Charged"], taxAdjustment, totalAdjustmentPerUnit )
		
		if abs(v1-v2) > 1e-5:
			logging.error("**** WARN: Failed Consistency Check %s %s %s %s %s %s", key, v1, v2, v1-v2, taxAdjustment, totalAdjustment)

 
	except KeyError as e:
		logging.warn("**** WARN: No Order Report Entry %s **** Shipping Soon?", key)
		value["Tax Adjustment"] = 0.0
		value["Total Adjustment"] = 0.0
		value["Total Adjustment Per Unit"] = 0.0



for row in reader:
	newLine = []

	if skip_row(row):
		logging.debug("**** Skipping Row ****")
		logging.debug(row)
		continue

	for nh in newHeader:
		#print "-----", nh

		if nh == "*ContactName":
			value = "Amazon"
		elif nh == "*InvoiceNumber":
			value = row["Order ID"]  
		elif nh == "*InvoiceDate" or nh == "*DueDate":
			value = row["Order Date"] 
		elif nh == "Description":
			value = row["Title"] 
		elif nh == "*Quantity":
			value = row["Quantity"] 
		elif nh == "*UnitAmount":
			#print "%%%", row["Item Total"].replace('$',''), row["Quantity"]
			if row["Quantity"] == "0":
				raise Exception("Quantity=0 row", row)
			value = str( float(row["Item Total"].replace('$','')) / float(row["Quantity"]) + orderBreakdown[key]["Total Adjustment Per Unit"] ) 
		elif nh == "*AccountCode":
			value = '' # set account code here
		elif nh == "*TaxType":
			value = 'Tax Exempt (0%)' 
		else:
			value = ''

		#print "###", nh, value
		newLine.append( value )

	#print ",".join(newLine)
	writer.writerow( newLine )

outfile.close()
logging.info( "*** DONE ***" )




