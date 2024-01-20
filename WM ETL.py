import glob, os
import openpyxl
import pandas as pd
import numpy as np
import sqlite3 as db


def getDateRow():
	#check the 3 new mg items shaver 9000, s5466, bt515, scd306, joanne's new online gift sets
	# check aaron's new ohc skus
	# remove columns: returns from retail link pull
	# remove J
#31011927

	# excel raw file
	os.chdir(os.getcwd())

	rawFile = glob.glob("book*.xlsx")[0]
	df = pd.read_excel(rawFile)

	# get the name of the dummy column
	col_name = str(df.columns[1])

	#get the year and week string
	yearWeek = df[df[col_name].str.contains('Time Range',case=False,na=False)].iat[0,1].split(" ")[-1]
	year = int(yearWeek[0:4])
	week = int(yearWeek[4:6].lstrip('0'))

	# find out how many useless rows to skip
	skipRow = df[df[col_name] == 'Item Flags'].index[0] + 1

	cleaner(rawFile, skipRow, year, week)

	removeFile(rawFile)

def cleaner(filePath, skip, year, week):

	df = pd.read_excel(filePath, skiprows=skip, usecols=lambda col: col not in ["Item Flags"], engine='openpyxl')

	df.columns = df.columns.str.replace(' ', '')
	df['ECOMM'] = np.where(df['StoreName'].str.contains("ECOMM"), "Online", "In Store")
	df['Year'] = year
	df['Week'] = week
	df['ZipCode'] = df['ZipCode'].str.replace(' ', '')
	df['StoreName'] = df['StoreName'].str.upper()
	df['FinelineDescription'] = df['FinelineDescription'].str.upper()
	df['Category'] = df['VendorStkNbr'].apply(category)

	# loading the cleaned data into the database
	conn = db.connect('WM.db')
	df.to_sql(name='wmPOS', con=conn,if_exists='append', index=False)

	query = """SELECT ItemNbr, VendorStkNbr, ItemDesc1, FinelineDescription, StoreNbr, StoreName,
		SUM(POSSales) AS posSales, SUM(POSQty) AS posQty, SUM(CurrStrOnHandQty) AS onHand, ECOMM, 
		Category, Week, Year FROM wmPOS WHERE Week=%d AND Year=%d AND 
		ItemNbr IN (SELECT ItemNbr FROM wmPOS GROUP BY ItemNbr HAVING SUM(POSQty) > 11) 
		GROUP BY VendorStkNbr, ECOMM, Week, Year""" %(week, year)

	query2 = """SELECT ItemNbr, VendorStkNbr, FinelineDescription, Category, 
		ROUND(CAST(COUNT(CASE WHEN CurrStrOnHandQty > 0 THEN 1 END) AS REAL)/ CAST(COUNT(CurrStrOnHandQty) AS REAL), 2) AS Instock, Week, Year
		FROM wmPOS 
		WHERE Week=%d AND Year=%d AND 
		ItemNbr IN (SELECT ItemNbr FROM wmPOS GROUP BY ItemNbr HAVING SUM(POSQty) > 11)
		GROUP BY ItemNbr, Week, Year"""

	df = pd.read_sql_query(query, conn)
	df2 = pd.read_sql_query(query2, conn)

	conn.close()

	#change header to True for first run
	df.to_csv('wmPOS.csv', mode='a', index=False, header=False)
	df2.to_csv('instock.csv', mode='a', index=False, header=False)

def removeFile(filePath):

	if os.path.isfile(filePath):
	    os.remove(filePath)
	else:    ## Show an error ##
	    print("Error: %s file not found" % filePath)

def category(sku):
	
	result = ""

	categoryDict = {
		"SCY":"MCC", "TRA":"MCC", "SCD":"MCC", "SCF":"MCC",
		"BRT":"Beauty", "BRE":"Beauty", "BRL":"Beauty", "BRP":"Beauty",
		"CC":"Male Grooming", "HQ":"Male Grooming", "AT":"Male Grooming", "BG":"Male Grooming",
		"BT":"Male Grooming", "MG":"Male Grooming", "QP":"Male Grooming", "NT":"Male Grooming",
		"HC":"Male Grooming", "SP":"Male Grooming", "SH":"Male Grooming", "RQ":"Male Grooming",
		"HX":"OHC", "HY":"OHC", "BH":"OHC", "HP":"Beauty",
		"S":"Male Grooming", "DIS":"OHC", "HF":"HSS"
	}

	if sku[0:3].upper() in categoryDict:
		result = categoryDict[sku[0:3]]

	elif sku[0:2].upper() in categoryDict:
		result = categoryDict[sku[0:2]]

	elif sku[0].upper() in categoryDict:
		result = categoryDict[sku[0]]

	return result


#SELECT Count(ItemNbr) FROM wmPOS; 47982

if __name__ == "__main__":
	getDateRow()

