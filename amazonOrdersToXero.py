#!/usr/bin/env python
# 
# Converts an Amazon order history download into a CSV to import into Xero
# Generate Amazon reports (Items & Orders): https://www.amazon.com/gp/b2b/reports
# Import into Xero:  https://go.xero.com/Accounts/Payable/Dashboard/
#
# Amazon Item and Order reports are combined to produce a CSV which can be imported 
# into Xero.  Both Amazon reports are necessary because the item report does not
# contain all necessary information.  Furthermore, some data is inconsistent between
# the two reports (e.g. taxes).
#

import sys
import csv
import logging

logging.basicConfig(level=logging.INFO)

def parse_key(row):
	'''Extract the unique key-tuple for a row'''
	return (row["Order ID"], row["Carrier Name & Tracking Number"])

def skip_row(row):
	'''Should this row in the Amazon Item Report be skipped'''
	return row["Order Status"] == "Cancelled" and row["Quantity"] == "0"

class OrderReport:
	'''Data from the Amazon Order Report'''

	def __init__(self, fileName):
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


class OrderAdjustments:
	'''Adjustments that need to be made to the Amazon Item Report.  The adjustments are determined from data in the
	Amazon Order Report which is not present in the Amazon Item Report.'''

	def __init__(self, orderReport, fileNameItem):
		itemReader = csv.DictReader(open(fileNameItem), delimiter=',', quotechar='"')
		self._data = {}

		for row in itemReader:
			if skip_row(row):
				logging.debug("**** Skipping Row ****")
				logging.debug(row)
				continue

			key = parse_key(row)

			if key not in self._data:
				value = {"Quantity":0, "Item Total": 0.0, "Shipping soon": 0, "Item Subtotal Tax": 0.0}
				self._data[key] = value

			value = self._data[key]
			value["Quantity"] += float(row["Quantity"])
			value["Item Total"] += float(row["Item Total"].replace('$',''))
			value["Item Subtotal Tax"] += float(row["Item Subtotal Tax"].replace('$',''))
	
			if row["Order Status"] == "Shipping soon":
				value["Shipping soon"] += 1

	
		for key in self._data:
			value = self._data[key]
			logging.debug("***************************************************")
			logging.debug("%s %s", key, value)

			try:
				orv = orderReport[key]
				taxAdjustment = float(orv["Tax Charged"]) - value["Item Subtotal Tax"] 
				totalAdjustment = float(orv["Shipping Charge"]) - float(orv["Total Promotions"]) + taxAdjustment
				totalAdjustmentPerUnit = totalAdjustment / value["Quantity"]
		
				value["Tax Adjustment"] = taxAdjustment
				value["Total Adjustment"] = totalAdjustment
				value["Total Adjustment Per Unit"] = totalAdjustmentPerUnit

				v1 = float(value["Item Total"]) + totalAdjustment
				v2 = float(orv["Total Charged"])
				logging.debug("Consistency Check: %s %s %s %s %s %s %s %s", key, v1, v2, abs(v1-v2)<1e-5, \
					value["Item Subtotal Tax"], orv["Tax Charged"], taxAdjustment, totalAdjustmentPerUnit )
		
				if abs(v1-v2) > 1e-5:
					logging.critical("Failed Consistency Check %s %s %s %s %s %s", key, v1, v2, v1-v2, taxAdjustment, totalAdjustment)

 
			except KeyError as e:
				# Handle cases where the order report does not contain any information for the key.
				logging.error("No Order Report Entry %s **** Shipping Soon?", key)
				value["Tax Adjustment"] = 0.0
				value["Total Adjustment"] = 0.0
				value["Total Adjustment Per Unit"] = 0.0


	def adjustment_per_unit(self, key):
		'''Returns the amount that the price of each unit should be adjusted by.  This is the total adjustment divided by the quantity in the order.'''
		return self._data[key]["Total Adjustment Per Unit"]


def write_xero_file(fileNameItem, fileNameOutput, orderAdjustments):
	'''Create a xero import file from the Amazon Order and Item Reports.'''

	reader = csv.DictReader(open(fileNameItem), delimiter=',', quotechar='"')
	outfile = open(fileNameOutput,'wb') 
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
		"Currency",
#		"AmazonAdjustment"	
		]

	writer.writerow( newHeader )

	for row in reader:
		newLine = []

		if skip_row(row):
			logging.debug("**** Skipping Row ****")
			logging.debug(row)
			continue

		key = parse_key(row)

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

				if abs(orderAdjustments.adjustment_per_unit( key )) > 1e-5:
					logging.warn("Adjusted: %s %s", key, orderAdjustments.adjustment_per_unit( key ))

				value = str( float(row["Item Total"].replace('$','')) / float(row["Quantity"]) + orderAdjustments.adjustment_per_unit( key ) )
			elif nh == "*AccountCode":
				value = '' # set account code here
			elif nh == "*TaxType":
				value = 'Tax Exempt (0%)'
			elif nh == "AmazonAdjustment":
#				if abs(orderAdjustments.adjustment_per_unit( key )) > 1e-5:
#					logging.warn("Adjusted: %s %s", key, orderAdjustments.adjustment_per_unit( key ))
#
				value = str(orderAdjustments.adjustment_per_unit( key )) 
			else:
				value = ''

			#print "###", nh, value
			newLine.append( value )

		#print ",".join(newLine)
		writer.writerow( newLine )

	outfile.close()


if len(sys.argv) !=3:
	print "Usage: <amzn_item_report.csv> <amzn_order_report.csv>"
	sys.exit(-1)

fileNameItem = sys.argv[1]
fileNameOrder = sys.argv[2]
fileNameOutput = fileNameItem.replace('.csv','.XERO.csv')
logging.info("fileNameItem: %s", fileNameItem)
logging.info("fileNameOrder: %s", fileNameOrder)
logging.warn("fileNameOutput: %s", fileNameOutput)

orderReport = OrderReport(fileNameOrder)
orderAdjustments = OrderAdjustments(orderReport, fileNameItem)

write_xero_file(fileNameItem, fileNameOutput, orderAdjustments)

logging.info( "*** DONE ***" )




