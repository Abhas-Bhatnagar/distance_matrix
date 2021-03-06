from xlrd import open_workbook
from xlwt import Workbook
import json,urllib.request
import math
import xlsxwriter
workbook = xlsxwriter.Workbook('output.xlsx')										#NAME OF GENERATED FILE
worksheet = workbook.add_worksheet()
wb = open_workbook('input.xlsx')														#NAME OF INPUT FILE CONTAINING ONE SHEET
wb_sheet_name="Sheet1"																	#NAME OF WORKSHEET WITHIN WORKBOOK

def radial(lat1,lon1,lat2,lon2):														#CALCULATE DISTANCE IN km
	radius = 6371 # km
	dlat=math.radians(lat2-lat1)
	dlon=math.radians(lon2-lon1)
	a = math.sin(dlat/2)*math.sin(dlat/2)+ math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dlon/2)*math.sin(dlon/2)
	c = 2*math.atan2(math.sqrt(a),math.sqrt(1-a))
	d = radius * c
	return(d)
	
		
def distance(lat_o,lon_o,lat_d,lon_d,row_o,col_d):
	url="http://dev.virtualearth.net/REST/V1/Routes/Walking?wp.0="+str(lat_o)+","+str(lon_o)+"&wp.1="+str(lat_d)+","+str(lon_d)+"&optmz=distance&output=json&key=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
	#IN ABOVE URL STRING FOLLWING CAN BE CHANGED TO		#mode=walking 				#units=KM                           											                #key=XXXX
	#FOR MORE REFERENCE #https://msdn.microsoft.com/en-us/library/gg636955.aspx https://msdn.microsoft.com/en-us/library/ff859477.aspx
	r_distance=radial(lat_o,lon_o,lat_d,lon_d)											#RETURNS DISTANCE IN KM
	if lat_o==lat_d and lon_o==lon_d:
		worksheet.write(row_o,col_d,"0")
		print("0")
	elif r_distance<=1:																	#NEED TO BE CHANGED HERE 1 DENOTES 1km
		data = urllib.request.urlopen(url).read()
		data=data.decode('utf-8')
		d = json.loads(data)
		print(d["resourceSets"][0]["resources"][0]["travelDistance"]*1000)
		worksheet.write(row_o,col_d,d["resourceSets"][0]["resources"][0]["travelDistance"]*1000)
	else:
		worksheet.write(row_o,col_d,"out of range")		
		#print("out of range")
		
		
def vertical(lat_o,lon_o,row_o):														#DO ENTRY IN COLUMN WISE
	for s in wb.sheets():
		lo=la=s.ncols
		if(s.name==wb_sheet_name):
			for row in range(s.nrows):
				for col in range(s.ncols):
					value  = (s.cell(row,col).value)
					if value=="Latitude":
						la=col
					elif value=="Longitude":
						lo=col
					elif la==col:
						lat_d=value
					elif lo==col:
						lon_d=value
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
					vertical(lat,value,row)													#CALLED NO. OF SCHOOL TIMES w.r.t ONE SCHOOL
				elif SCHCD_row==col:
					worksheet.write(0,row,value)
					worksheet.write(row,0,value)
		worksheet.write(0,0,"SCHCD")
		workbook.close()
