from xlrd import open_workbook
from xlwt import Workbook
import json
import urllib2
import math
import xlsxwriter
import sys
import time

workbook = xlsxwriter.Workbook('output1.xlsx')										#NAME OF GENERATED FILE
worksheet = workbook.add_worksheet()
wb = open_workbook('Nellore_UP_6_8.xlsx')#SKLM_Telugu_new.xlsx														#NAME OF INPUT FILE CONTAINING ONE SHEET
wb_sheet_name="total_enrolment"																	#NAME OF WORKSHEET WITHIN WORKBOOK


def radial(lat1,lon1,lat2,lon2):														#CALCULATE DISTANCE IN km
	radius = 6371 # km
	dlat=math.radians(lat2-lat1)
	dlon=math.radians(lon2-lon1)
	a = math.sin(dlat/2)*math.sin(dlat/2)+ math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dlon/2)*math.sin(dlon/2)
	c = 2*math.atan2(math.sqrt(a),math.sqrt(1-a))
	d = radius * c
	return(d)
		
def distance(lat_o,lon_o,lat_d,lon_d,row_o,col_d):
	url="https://maps.googleapis.com/maps/api/distancematrix/json?units=metric&origins="+str(lat_o)+","+str(lon_o)+"&destinations="+str(lat_d)+","+str(lon_d)+"&mode=walking&key=XXXXXXXXXXXXXXXXXXXXXXX"
	#IN ABOVE URL STRING FOLLWING CAN BE CHANGED TO		#mode=walking 				#units=KM                           											                #key=XXXX
	r_distance=radial(lat_o,lon_o,lat_d,lon_d)											#RETURNS DISTANCE IN KM
	if lat_o==lat_d and lon_o==lon_d:
		worksheet.write(row_o,col_d,"0")
		#if ((s.cell(row_o,0).value) % 4 == 0 ): 
			#print(str((s.cell(row_o,0).value)))
	elif r_distance<=4:																	#NEED TO BE CHANGED HERE 1 DENOTES 1km
			data = urllib2.urlopen(url).read()
			#time.sleep(1)
			data=data.decode('utf-8')
			d = json.loads(data)
			#worksheet.write(row_o,col_d,d["resourceSets"][0]["resources"][0]["travelDistance"]*1000)
			#print(d["rows"][0]["elements"][0]["distance"]["value"])
			worksheet.write(row_o,col_d,d["rows"][0]["elements"][0]["distance"]["value"])
	else:
		worksheet.write(row_o,col_d,"NA")
		#worksheet.write(col_d,row_o,"NA")
		
		
		
def vertical(lat_o,lon_o,row_o):														#DO ENTRY IN COLUMN WISE
	for s in wb.sheets():
		lo=la=set=s.ncols
		if(s.name==wb_sheet_name):
			for row in range(s.nrows):
				for col in range(s.ncols):
					value  = (s.cell(row,col).value)
					if value=="Latitude":
						la=col
					elif value=="Longitude":
						lo=col
					elif value=="SET":
						set=col
					elif la==col and (s.cell(row_o,set).value)==(s.cell(row,set).value):
						lat_d=value
					elif lo==col and (s.cell(row_o,set).value)==(s.cell(row,set).value):
							lon_d=value
							#print(str((s.cell(row,set).value)))
							distance(lat_o,lon_o,lat_d,lon_d,row_o,row)
											
						
					
#starting point																			#DO ENTRY IN ROW WISE
for s in wb.sheets():
	lo=la=s.ncols
	SCHCD_row=s.ncols
	if(s.name==wb_sheet_name):
		for row in range(s.nrows):
			for col in range(s.ncols):
				value  = (s.cell(row,col).value)
				if value=="Latitude":														#SHOULD BE SAME AS COLUMN NAME OF LATITUDE & IS CASE SENSITIVE i.e UPPER CASE CHARACTER IS DIFFERENT FROM LOWER CASE CHARACTER
					la=col
				elif value=="Longitude":													#SHOULD BE SAME AS COLUMN NAME OF LONGITUDE & IS CASE SENSITIVE
					lo=col
				elif value=="SCHCD":														#SHOULD BE SAME AS COLUMN NAME OF SCHOOL CODE & IS CASE SENSITIVE
					SCHCD_row=col
				elif la==col:
					lat=value
				elif lo==col:
					if (row %20 ==0):
						print(row)
						#time.sleep(1)
					vertical(lat,value,row)													#CALLED NO. OF SCHOOL TIMES w.r.t ONE SCHOOL
				elif SCHCD_row==col:
					worksheet.write(0,row,value)
					worksheet.write(row,0,value)
		worksheet.write(0,0,"SCHCD")
		workbook.close()
