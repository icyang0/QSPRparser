# ####todo
# implement file name selection
# for the lines selecing specific column, make it not hard coded. ie make it so we can choose whatever column based on the selected paramterers


from openpyxl import load_workbook
from openpyxl import Workbook
import numpy as np

INPUTFILE1 = 'test-orig1.xlsx'
OUTPUTFILE1 = 'edited1-test1.xlsx'
INPUTFILE2 = 'test-orig2.xlsx'
OUTPUTFILE2 = 'edited1-test2.xlsx'


columnTrans = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC')
#create the tuple of attributes to look for
keyParametersTX = ('ParameterName', 'channel', 'txChainMask', 'rate', 'powerLevel', 'NumValue')
keyParametersRX = ('ParameterName', 'channel', 'txChainMask', 'rate', 'powerLevel', 'NumValue')

keyParametersTXrow = ('evm', 'avgTxPower')

def colNumToColStr(numberOfCol):
	return columnTrans[numberOfCol + 1]

def indexNumToColStr(index):
	return columnTrans[index]

def parseFile(rawFileName, newSaveFileName):
	#load the file
	wb = load_workbook(rawFileName)
	ws = wb.active
	#remove the first 3 rows
	ws.delete_rows(0,3)


	#find the column index of all the key paramters and creates a tuple with them
	headerRow = ws[1]
	keyParameterIndexTup = ()
	for i in range(len(keyParametersTX)):
		for x in range(len(headerRow)):
			if (keyParametersTX[i] in headerRow[x].value):
				dog = (x,)
				keyParameterIndexTup = keyParameterIndexTup + dog
	
	#go through and delete the columns that are not in the key parameters			
	i = len(keyParameterIndexTup) - 1
	for z in range( len(headerRow) -1, -1, -1):
		#keyParameterIndexTup[i]
		if (keyParameterIndexTup[i] != z):
			ws.delete_cols(z + 1, 1)
		else:
			i = i - 1


		
	###################remove the rows not of interest#########################


	#Create a list of the testnames
	parameterNameCol = ws[indexNumToColStr(keyParametersTX.index('ParameterName'))]
	parameterNameColList = []
	for m in range(len(parameterNameCol)):
		parameterNameColList.append(parameterNameCol[m].value)
	
	#create a list of the index of the rows of the test names	
	keyParameterIndexCol = []
	for i in range(len(keyParametersTXrow)):
		for x in range(len(parameterNameColList)):
			if (keyParametersTXrow[i] == parameterNameColList[x]):
				keyParameterIndexCol.append(x)

	keyParameterIndexCol.sort()
	#print (keyParameterIndexCol)	
	
	#go through and delete the columns that are not in the key parameters of tests (in backwards order)
	i = len(keyParameterIndexCol) - 1
	for z in range( len(parameterNameCol) -1, -1, -1):
		#saves the column names (if you want to remove it, remove the z !=0 condition)
		if (keyParameterIndexCol[i] != z):
		#if ((keyParameterIndexCol[i] != z) and (z !=0)):
			ws.delete_rows(z + 1, 1)
		else:
			i = i - 1
		
		
	################### put the avgtxpower in same row as the evm ##############################
	#Create a list of the numValues
	numValueColTup = ws[indexNumToColStr(keyParametersTX.index('NumValue'))]
	evmColList = []
	avgTXPowerColList = []
	for m in range(len(numValueColTup)):
		#if we are on an even index, then it's an EVM value, so seperate that out
		if (m%2 == 0):
			evmColList.append(numValueColTup[m].value)
		else:
			avgTXPowerColList.append(numValueColTup[m].value)
		

	#insert a new column before the evm measuremenr
	ws.insert_cols(keyParametersTX.index('NumValue') + 1)
		
	#delete the avgtxpower rows, and add the avgtxpower values to the evm rows from the backwards direction
	for z in range(len(numValueColTup) -1, -1, -1):
		#if on an odd index (i.e. we are on a avgtxpower) then delete it
		if (z%2 != 0):
			ws.delete_rows(z+1, 1)
		#if on an even index (i.e. we are on evm row) then keep it and append the avgtxpower onto it
		else:
			ws[(indexNumToColStr(keyParametersTX.index('NumValue'))) + (str(z+1))] = avgTXPowerColList[int(z/2)]
	
	#save the file!	
	wb.save(newSaveFileName)	
		
	return 1	


###########################save the test conditions#####################
def combineFile(file1, file2):
	testCondsTup = ws[('B') + ':' + ('D')]
	testConds = (np.array(list(testCondsTup))).T

	newDataIndexList = [0]

	#loop through the saved data and see if there		
	for irow in range(1, len(testConds)):
		for icol in range(len(testConds[0])):
			#print (testConds[irow][icol].value)
			prevVal = testConds[irow-1][icol].value
			if ((testConds[irow][icol].value != prevVal) & (newDataIndexList[len(newDataIndexList) - 1] != irow)):
				newDataIndexList.append(irow)
			
			
	print (newDataIndexList)
	
	
	
	return 1
	
		
		

parseFile(INPUTFILE1, OUTPUTFILE1)
parseFile(INPUTFILE2, OUTPUTFILE2)
	
	
	
