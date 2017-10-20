import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import csv
import os
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl
import time
import datetime
from Tkinter import Tk
from tkFileDialog import askopenfilename, askdirectory

'''
*copy below and type %paste in IPython console

%cd C:\Users\S3Research\Desktop\Python Example
%run PLEXOSprocessing.py

'''
# Following should be updated
refFileFolder = "C:\Users\S3Research\Desktop\Python Example"
refFile = "ref_pg.xlsx"  # Reference File Folder
desktop = "C:\Users\S3Research\Desktop"	# So program can save output file to desktop

# Following can be changed
toXLSX = True  # Outputs case_output.xlsx if true, else outputs file tabs as .csv's
afternoon_hours = [11, 12, 13, 14, 15, 16]  # For afternoon SOx calculations
compare = True  # Produces the comparison_Percent .csv's

startTime = time.time()		# To time the program

Tk().withdraw()		# Weird box pops up if you don't have this
outputFolderName = raw_input('Enter a name for output folder: ')  # Prompts user for name for output folder
outputFolder = desktop + "/" +outputFolderName  # Output folder is going to be saved to desktop
os.makedirs(outputFolder)	# Creates output folder

number_of_cases = int(raw_input('How many cases: '))	# prompts user and saves input
cases = []  # Create empty case array: [['case1',folderName], ['case2',folderName]]

# Get name and folder for all other cases and store in cases 
for i in range(0,number_of_cases):
	s = str(i)
	temp_name = raw_input('Enter case number '+s+' name: ')
	temp_dir = askdirectory(initialdir="C:\Users\ktsanders\Documents\PLEXOS\PLEXOS_ERCOT_MODEL\\",
							title='Please select a directory for this case:')
	cases.append([temp_name, temp_dir])
	
ind = ['SOx', 'NOx', 'CO2', 'WC', 'WW', 'Afternoon SOx', 'Coal', 'NG', '%Coal', '%NG', '1','2','3','4','5','6','7','8','9','10','11','12']
clmns = []
for case in enumerate(cases):
	clmns.append(case[1][0])
	
dfOverall = pd.DataFrame(index=ind,columns=clmns)
	
print 'Finding Reference File'
cwd = os.getcwd()
os.chdir(refFileFolder)
xls = pd.ExcelFile(refFile)		# Get reference file 
dfRef = xls.parse('water_use_definitions')  # store first sheet as DataFrame ex: dfRef[dfRef['class']=='Generator']
dfCNums = xls.parse('ClassNums') 	# store second sheet as DataFrame
del dfRef['id']		# the index is the id so no need
numGenerators = dfRef.index.size # Number of generators to loop through; actually +1 but no 0 index so range(1,numGenerators) is correct
os.chdir(cwd)

# Main for loop that loops through cases
j = 0
for c in cases:
	j += 1
	path = c[1] + '\interval'  # Folder that PLEXOS outputs to
	case = c[0]  # Name of case
	print 'Starting case: ' + str(case)

	caseFile = outputFolder+"\Analysis_"+case
	os.makedirs(caseFile)
	
	# df holds all the generation and is created with first gen file
	df = pd.read_csv(path+'\ST Generator(1).Generation.csv', sep=',', index_col='DATETIME')
	v1 = df.index.size  #variable 1: number of rows or datetimes

	n=0
	print '   Checking generation files'
	# Loop through all generators and simply see if file exists or create new empty file
	for i in range(1,numGenerators):
		try:
			files = pd.read_csv(path+'\ST Generator('+ str(i) + ').Generation.csv', sep=',')
		except:
			print str(i) + ' Not Exist'
			n += 1
			newfile = pd.DataFrame({'DATETIME' : df.index,
									'VALUE' : pd.Series([np.nan] * v1)})
			newfile.to_csv(path + '\ST Generator('+ str(i) + ').Generation.csv')
 
	print '   There are ' + str(n) + ' files that not exist!'

	print '   Creating generation sheet'
	df.columns = [1]  # Rename first column 1 for generator 1 (originally reads 'VALUE' from .csv)
	for i in range(2,numGenerators):  # Read all gen files
		file = pd.read_csv(path+'\ST Generator('+ str(i) + ').Generation.csv', sep=',', index_col='DATETIME')
		df[i] = file['VALUE']  # Populate df with generation values; column title is gen number

	arr1 = []  # Array that holds the file names for fuel offtake
	arr1.append(np.nan)  # want index to represent gen number so index 0 is set empty
	arr_wc = []		# Array that holds water consumption files name (Emissions(4).Generator.Production)
	arr_wc.append(np.nan)	# index 0 set empty
	arr_ww = []		# Array that holds water withdraw file name (Emissions(5).Generator.Production)
	arr_ww.append(np.nan)	# index 0 set empty
	dfFuelOff = pd.DataFrame({1:pd.Series([np.nan]*df.index.size)})		# Create fuel offtake df
	dfFuelOff.index = df.index		# Set index as dates
	print '   Finding offtake and water files'	 
	
	# loop through all generators
	for i in range(1,400):
		count = 1
		# Loop through fuel types to find correct file name and store in arr1
		for f in range(1,6):
			try:
				file = path + '\ST Generator('+ str(i) + ').Fuels('+ str(f) + ').Offtake.csv'
				dftemp = pd.read_csv(file)
				arr1.append(str(file))
				break
			except:
				if count == 5:
					arr1.append(np.nan)
				count += 1
		
		# See if WC production file exists and store file name in arr_wc
		try:
			wcFile = path + '\ST Emission(4).Generators('+ str(i) + ').Production.csv'
			dftemp_wc = pd.read_csv(wcFile)
			arr_wc.append(str(wcFile))
		# If a file doesn't, store no value in arr_wc so that index remains same as generator number
		except:
			arr_wc.append(np.nan)
		
		# See if WW production file exists and store file name in arr_wc
		try:
			wwFile = path + '\ST Emission(5).Generators('+ str(i) + ').Production.csv'
			dftemp_ww = pd.read_csv(wwFile)
			arr_ww.append(str(wwFile))
		# If a file doesn't, store no value in arr_ww so that index remains same as generator number
		except:
			arr_ww.append(np.nan)
			
	size = dfFuelOff.index.size
	print '   Loading/translating offtake files'
	for i in range(1,len(arr1)):
		if type(arr1[i]) is str:  # as opposed to nan
			dfFuelOff[i] = pd.read_csv(arr1[i], index_col=0)['VALUE']
		else:
			dfFuelOff[i] = pd.Series([np.nan]*size)
		
	print '   Loading/translating water consumption files'
	df_wc = pd.DataFrame({1:pd.Series([np.nan]*df.index.size)})
	df_wc.index = df.index
	for i in range(1,len(arr_wc)):
		if type(arr_wc[i]) is str:
			df_wc[i] = pd.read_csv(arr_wc[i], index_col=0)['VALUE']
		else:
			df_wc[i] = pd.Series([np.nan]*size)
	df_wc = df_wc/8.33
		
	print '   Loading/translating water withdrawal files'
	df_ww = pd.DataFrame({1:pd.Series([np.nan]*df.index.size)})
	df_ww.index = df.index
	for i in range(1,len(arr_ww)):
		if type(arr_ww[i]) is str:
			df_ww[i] = pd.read_csv(arr_ww[i], index_col=0)['VALUE']
		else:
			df_ww[i] = pd.Series([np.nan]*size)
	df_ww = df_ww/8.33
		
	dfSOx = pd.DataFrame({1 : dfFuelOff[1]*dfRef.loc[1]['SOx']})
	dfNOx = pd.DataFrame({1 : dfFuelOff[1]*dfRef.loc[1]['NOx']})
	dfCO2 = pd.DataFrame({1 : dfFuelOff[1]*dfRef.loc[1]['CO2']})
	for i in range(1,dfRef.index.size):
		dfSOx[i] = dfFuelOff[i]*dfRef.loc[i]['SOx']
		dfNOx[i] = dfFuelOff[i]*dfRef.loc[i]['NOx']
		dfCO2[i] = dfFuelOff[i]*dfRef.loc[i]['CO2']

	print '   Finding total data'
	emptySeries = pd.Series(dfCNums['WaterClassStr'].index.size*[np.nan])
	dfRef['TotalGen'] = df.sum(axis=0)
	dfTypeData = pd.DataFrame({'GenType' : dfCNums['WaterClassStr'],
								'Total Gen' : emptySeries,
								'Total SOx' : emptySeries,
								'Total NOx' : emptySeries,
								'Total CO2' : emptySeries}) # make *(1+number_of_classes)
	dfTypeData.reindex(columns=['GenType','Total Gen','Total SOx','Total NOx','Total CO2'])

	dfNoType = dfRef[dfRef['WaterClassNum'].isnull() & dfRef['TotalGen'].notnull()]
	arrTemp = []
	arrTemp.append(np.nan)
	print '   Agregating gen and emissions by cooling type'
	for i in range(1,dfTypeData.index.size):
		# Finds totals by type
		dfTypeData['Total Gen'][i] = dfRef[dfRef['WaterClassNum']==i]['TotalGen'].sum()
		dfTypeData['Total SOx'][i] = dfRef[dfRef['WaterClassNum']==i]['SOx'].sum()
		dfTypeData['Total NOx'][i] = dfRef[dfRef['WaterClassNum']==i]['NOx'].sum()
		dfTypeData['Total CO2'][i] = dfRef[dfRef['WaterClassNum']==i]['CO2'].sum()
		tempArr = dfRef[dfRef['WaterClassNum']==i].index
		arrTemp.append(tempArr)
		
		tempSeriesGen = df[tempArr].sum(axis=1)
		tempSeriesSOx = dfSOx[tempArr].sum(axis=1)
		tempSeriesNOx = dfNOx[tempArr].sum(axis=1)
		tempSeriesCO2 = dfCO2[tempArr].sum(axis=1)
		tempSeriesFO = dfFuelOff[tempArr].sum(axis=1)
		tempSeriesWC = df_wc[tempArr].sum(axis=1)
		tempSeriesWW = df_ww[tempArr].sum(axis=1)
		if i == 1:
			dfTGen = pd.DataFrame({1:tempSeriesGen})
			dfTSOx = pd.DataFrame({1:tempSeriesSOx})
			dfTNOx = pd.DataFrame({1:tempSeriesNOx})
			dfTCO2 = pd.DataFrame({1:tempSeriesCO2})
			dfTFO = pd.DataFrame({1:tempSeriesFO})
			dfTWC = pd.DataFrame({1:tempSeriesWC})
			dfTWW = pd.DataFrame({1:tempSeriesWW})
		else:
			dfTGen[i] = tempSeriesGen
			dfTSOx[i] = tempSeriesSOx
			dfTNOx[i] = tempSeriesNOx
			dfTCO2[i] = tempSeriesCO2
			dfTFO[i] = tempSeriesFO
			dfTWC[i] = tempSeriesWC
			dfTWW[i] = tempSeriesWW	

	# Change all DataFrame indexes to datetime format
	df.index = pd.to_datetime(df.index)
	df_wc.index = pd.to_datetime(df_wc.index)
	df_ww.index = pd.to_datetime(df_ww.index)
	dfNOx.index = pd.to_datetime(dfNOx.index)
	dfSOx.index = pd.to_datetime(dfSOx.index)
	dfCO2.index = pd.to_datetime(dfCO2.index)
	dfTGen.index = pd.to_datetime(dfTGen.index)
	dfTSOx.index = pd.to_datetime(dfTSOx.index)
	dfTNOx.index = pd.to_datetime(dfTNOx.index)
	dfTCO2.index = pd.to_datetime(dfTCO2.index)
	dfTWC.index = pd.to_datetime(dfTWC.index)
	dfTWW.index = pd.to_datetime(dfTWW.index)
	
	# Afternoon hours
	dfAfternoonSOx = dfTSOx[dfTSOx.index.hour == afternoon_hours[0]]   # Creates DataFrame with first hour
	for t in afternoon_hours[1:]:
		dfAfternoonSOx = dfAfternoonSOx.append(dfTSOx[dfTSOx.index.hour == t])
	dfAfternoonSOx = dfAfternoonSOx.sort_index()
	dfAfternoonSOx = dfAfternoonSOx.resample('D',how='sum')			

	dfOverall[case]['SOx'] = dfTSOx.sum().sum()
	dfOverall[case]['NOx'] = dfTNOx.sum().sum()
	dfOverall[case]['CO2'] = dfTCO2.sum().sum()
	dfOverall[case]['WC'] = dfTWC.sum().sum()
	dfOverall[case]['WW'] = dfTWW.sum().sum()
	dfOverall[case]['Afternoon SOx'] = dfAfternoonSOx.sum().sum()
	dfOverall[case]['Coal'] = dfTGen.sum()[2] + dfTGen.sum()[3]
	dfOverall[case]['NG'] = dfTGen.sum()[5] + dfTGen.sum()[6] + dfTGen.sum()[7] + dfTGen.sum()[8] + dfTGen.sum()[9]
	dfOverall[case]['%Coal'] = dfOverall[case]['Coal'] / dfTGen.sum().sum()
	dfOverall[case]['%NG'] = dfOverall[case]['NG'] / dfTGen.sum().sum()
	for index in ind[10:]:
		dfOverall[case][index] = dfTGen.sum()[int(index)]
		
	stats = ['Gen','SOx','NOx','CO2','WC','WW','Afternoon SOx']
	col = []
	col2 = []
	for c in cases:
		col2.append(c[0])
		for (n,v) in enumerate(stats):  # Creates DataFrame column heads ['case1 Gen', 'case1 SOx',.....,'case2 Gen', 'case2 SOx']
			col.append(c[0] + " " + v)
	
	if j == 1:
		dfReportDailySOx = pd.DataFrame(index=dfTSOx.resample('D',how='sum').sum(axis=1).index,columns=col2)
		dfReportDailyNOx = pd.DataFrame(index=dfTNOx.resample('D',how='sum').sum(axis=1).index,columns=col2)
		dfReportDailyCO2 = pd.DataFrame(index=dfTCO2.resample('D',how='sum').sum(axis=1).index,columns=col2)
		dfReportDailyWC = pd.DataFrame(index=dfTWC.resample('D',how='sum').sum(axis=1).index,columns=col2)
		
	dfReportDailySOx[case] = dfTSOx.resample('D',how='sum').sum(axis=1)
	dfReportDailyNOx[case] = dfTNOx.resample('D',how='sum').sum(axis=1)
	dfReportDailyCO2[case] = dfTCO2.resample('D',how='sum').sum(axis=1)
	dfReportDailyWC[case] = dfTWC.resample('D',how='sum').sum(axis=1)
	
	if compare == True:
		if j == 1:		
			baseData = {'gen':df, 'wc':df_wc, 'ww':df_ww, 'NOx':dfNOx, 'SOx':dfSOx, 'CO2':dfCO2, 'Afternoon SOx':dfAfternoonSOx,
						'TGen':dfTGen, 'TSOx':dfTSOx, 'TNOx':dfTNOx, 'TCO2':dfTCO2, 'TWC':dfTWC, 'TWW':dfTWW}			

			dfReportMonthly = pd.DataFrame(index=dfTGen.resample('M',how='sum').sum(axis=1).index,columns=col)
			dfReportDaily = pd.DataFrame(index=dfTGen.resample('D',how='sum').sum(axis=1).index,columns=col)
		
			dfReportMonthly[case + " Gen"] = dfTGen.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " SOx"] = dfTSOx.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " NOx"] = dfTNOx.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " CO2"] = dfTCO2.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " WC"] = dfTWC.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " WW"] = dfTWW.resample('M',how='sum').sum(axis=1)
			dfReportMonthly[case + " Afternoon SOx"] = dfAfternoonSOx.resample('M',how='sum').sum(axis=1)
	
			dfReportDaily[case + " Gen"] = dfTGen.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " SOx"] = dfTSOx.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " NOx"] = dfTNOx.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " CO2"] = dfTCO2.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " WC"] = dfTWC.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " WW"] = dfTWW.resample('D',how='sum').sum(axis=1)
			dfReportDaily[case + " Afternoon SOx"] = dfAfternoonSOx.resample('D',how='sum').sum(axis=1)
		
		else:
			#Reductions
			dfRSOx = baseData['TSOx'] - dfTSOx
			dfRNOx = baseData['TSOx'] - dfTNOx
			dfRCO2 = baseData['TCO2'] - dfTCO2
			dfRWC = baseData['TWC'] - dfTWC
			dfRWW = baseData['TWW'] - dfTWW
			#Daily Reductions
			dfDRSOx = dfRSOx.resample('D', how='sum')
			dfDRNOx = dfRNOx.resample('D', how='sum')
			dfDRCO2 = dfRCO2.resample('D', how='sum')
			dfDRWC = dfRWC.resample('D', how='sum')
			dfDRWW = dfRWW.resample('D', how='sum')
			#Diference
			dfDGen = dfTGen - baseData['TGen']
			dfDSOx = dfTSOx	- baseData['TSOx']
			dfDNOx = dfTNOx	- baseData['TNOx']
			dfDCO2 = dfTCO2	- baseData['TCO2']
			dfDWC = dfTWC	- baseData['TWC']
			dfDWW = dfTWW	- baseData['TWW']
		
			dfDRSOx.to_csv(caseFile+"\DailyReductions_SOx_"+case+".csv")
			dfDRNOx.to_csv(caseFile+"\DailyReductions_NOx_"+case+".csv")
			dfDRCO2.to_csv(caseFile+"\DailyReductions_CO2_"+case+".csv")
			dfDRWC.to_csv(caseFile+"\DailyReductions_WC_"+case+".csv")
			dfDRWW.to_csv(caseFile+"\DailyReductions_WW_"+case+".csv")
		
			dfDGen.to_csv(caseFile+"\DifferenceFromBaseline_Gen_"+case+".csv")
			dfDSOx.to_csv(caseFile+"\DifferenceFromBaseline_SOx_"+case+".csv")
			dfDNOx.to_csv(caseFile+"\DifferenceFromBaseline_NOx_"+case+".csv")
			dfDCO2.to_csv(caseFile+"\DifferenceFromBaseline_CO2_"+case+".csv")
			dfDWC.to_csv(caseFile+"\DifferenceFromBaseline_WC_"+case+".csv")
			dfDWW.to_csv(caseFile+"\DifferenceFromBaseline_WW_"+case+".csv")
		
			dfReportMonthly[case + " Gen"] = (dfTGen.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " Gen"]) / dfReportMonthly[cases[0][0] + " Gen"]
			dfReportMonthly[case + " SOx"] = (dfTSOx.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " SOx"]) / dfReportMonthly[cases[0][0] + " SOx"]
			dfReportMonthly[case + " NOx"] = (dfTNOx.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " NOx"]) / dfReportMonthly[cases[0][0] + " NOx"]
			dfReportMonthly[case + " CO2"] = (dfTCO2.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " CO2"]) / dfReportMonthly[cases[0][0] + " CO2"]
			dfReportMonthly[case + " WC"] = (dfTWC.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " WC"]) / dfReportMonthly[cases[0][0] + " WC"]
			dfReportMonthly[case + " WW"] = (dfTWW.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " WW"]) / dfReportMonthly[cases[0][0] + " WW"]
			dfReportMonthly[case + " Afternoon SOx"] = (dfAfternoonSOx.resample('M',how='sum').sum(axis=1) - dfReportMonthly[cases[0][0] + " Afternoon SOx"]) /  dfReportMonthly[cases[0][0] + " Afternoon SOx"]
		
			dfReportDaily[case + " Gen"] = (dfTGen.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " Gen"]) / dfReportDaily[cases[0][0] + " Gen"]
			dfReportDaily[case + " SOx"] = (dfTSOx.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " SOx"]) / dfReportDaily[cases[0][0] + " SOx"]
			dfReportDaily[case + " NOx"] = (dfTNOx.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " NOx"]) / dfReportDaily[cases[0][0] + " NOx"]
			dfReportDaily[case + " CO2"] = (dfTCO2.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " CO2"]) / dfReportDaily[cases[0][0] + " CO2"]
			dfReportDaily[case + " WC"] = (dfTWC.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " WC"]) / dfReportDaily[cases[0][0] + " WC"]
			dfReportDaily[case + " WW"] = (dfTWW.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " WW"]) / dfReportDaily[cases[0][0] + " WW"]
			dfReportDaily[case + " Afternoon SOx"] = (dfAfternoonSOx.resample('D',how='sum').sum(axis=1) - dfReportDaily[cases[0][0] + " Afternoon SOx"]) / dfReportDaily[cases[0][0] + " Afternoon SOx"]
	
	
	if toXLSX == True:
		tSave = []
		print '   Creating workbook'
		t0 = time.time()
		writer = ExcelWriter(caseFile + "/" + case + '_Output.xlsx') # *& Same address
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[0]))

		print '     Writing: 1/5'
		t0 = time.time()
		dfTGen.to_excel(writer,'Gen By Type')
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[1]))	

		print '     Writing: 2/5'
		t0 = time.time()
		dfTCO2.to_excel(writer,'CO2 By Type')
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[2]))

		print '     Writing: 3/5'
		t0 = time.time()
		dfTNOx.to_excel(writer,'NOx By Type')
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[3]))

		print '     Writing: 4/5'
		t0 = time.time()
		dfTSOx.to_excel(writer,'SOx By Type')
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[4]))

		print '     Writing: 5/5'
		t0 = time.time()
		dfTWC.to_excel(writer,'WC by type')
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[5]))

		writer.save()
		tSave.append(time.time()-t0)
		print '      Time: ' + str(datetime.timedelta(seconds=tSave[6]))
	else:
		writer = ExcelWriter(caseFile + "/" + case + '_Output.xlsx') # *& Same address
		dfTGen.to_csv(caseFile + '\Gen by Type_' + case + '.csv')
		dfTCO2.to_csv(caseFile + '\CO2 by Type_' + case + '.csv')
		dfTNOx.to_csv(caseFile + '\NOx by Type_' + case + '.csv')
		dfTSOx.to_csv(caseFile + '\SOx by Type_' + case + '.csv')
		dfTWC.to_csv(caseFile + '\WC by Type_' + case + '.csv')
		
	df.to_csv(caseFile + '\Gen by Generator_' + case + '.csv')
	dfFuelOff.to_csv(caseFile + '\Fuel Offtake by Generator_' + case + '.csv')
	dfSOx.to_csv(caseFile + '\SOx by Generator_' + case + '.csv')
	dfNOx.to_csv(caseFile + '\NOx by Generator_' + case + '.csv')
	dfCO2.to_csv(caseFile + '\CO2 by Generator_' + case + '.csv')
	dfTFO.to_csv(caseFile + '\Fuel Offtake by Type_' + case + '.csv')
	dfTWW.to_csv(caseFile + '\WW by Type_' + case + '.csv')
	# dfTypeData.to_csv(caseFile + '\Total Gen by Type_' + case + '.csv')
	
	dfAfternoonSOx.to_csv(caseFile+"\AfternoonSOx_"+case+".csv")

if compare == True:
	dfReportMonthly.to_csv(outputFolder+"\Comparison_Percent_Monthly.csv")
	dfReportDaily.to_csv(outputFolder+"\Comparison_Percent_Daily.csv")
dfOverall = dfOverall.T
dfOverall.to_csv(outputFolder+"\Case_Data.csv")

dfReportDailySOx.to_csv(outputFolder+"\SOx_Daily_Data.csv")
dfReportDailyNOx.to_csv(outputFolder+"\NOx_Daily_Data.csv")
dfReportDailyCO2.to_csv(outputFolder+"\CO2_Daily_Data.csv")
dfReportDailyWC.to_csv(outputFolder+"\WC_Daily_Data.csv")

	
stopTime = time.time()
totalTime = stopTime - startTime
print 'Run Time: '+ str(datetime.timedelta(seconds=totalTime))

