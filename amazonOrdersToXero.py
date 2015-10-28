#!/usr/bin/env python
# 
# Converts an Amazon order history download into a CSV to import into Xero
# Generate Amazon report: https://www.amazon.com/gp/b2b/reports
#

import sys
import csv

if len(sys.argv) !=2:
	print "Usage <amznfile.csv>"
	sys.exit(-1)

fileName = sys.argv[1]
print "Output: ", fileName

reader = csv.DictReader(open(fileName), delimiter=',', quotechar='"')
outfile = open(fileName.replace('.csv','.NEW.csv'),'wb')
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

for row in reader:
	newLine = []

	if row["Order Status"] == "Cancelled" and row["Quantity"] == "0":
		print "**** Skipping Row ****"
		print row
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
				print row
			value = str( float(row["Item Total"].replace('$','')) / float(row["Quantity"]) ) 
		elif nh == "*AccountCode":
			value = '' #todo set account code
		elif nh == "*TaxType":
			value = 'Tax Exempt (0%)' #'INPUT' #todo not sure
		else:
			value = ''

		#print "###", nh, value
		newLine.append( value )

	#print ",".join(newLine)
	writer.writerow( newLine )
#	"EmailAddress",
#	"POAddressLine1",
#	"POAddressLine2",
#	"POAddressLine3",
#	"POAddressLine4",
#	"POCity",
#	"PORegion",
#	"POPostalCode",
#	"POCountry",
#	"InventoryItemCode",
#	"Description",
#	"TrackingName1",
#	"TrackingOption1",
#	"TrackingName2",
#	"TrackingOption2",
#	"Currency"	

outfile.close()
print "*** DONE ***"




