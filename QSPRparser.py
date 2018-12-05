from openpyxl import load_workbook
from openpyxl import Workbook

columnTrans = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC')
#create the tuple of attributes to look for
keyParametersTX = ('ParameterName', 'channel', 'txChainMask', 'rate', 'powerLevel', 'NumValue')
keyParametersRX = ('ParameterName', 'channel', 'txChainMask', 'rate', 'powerLevel', 'NumValue')

keyParametersTXrow = ('evm', 'avgTxPower')


wb = load_workbook(filename = 'test.xlsx')
ws = wb.active



#########these format the original, so remove if youre testing on the formatted version #################
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


#remove the rows not of interest
parameterNameCol = ws['A']
keyParameterIndexTupCol = ()
for i in range(len(keyParametersTXrow)):
	for x in range(len(parameterNameCol)):
		if (keyParametersTXrow[i] in parameterNameCol[x].value):
			dog = (x,)
			keyParameterIndexTupCol = keyParameterIndexTupCol + dog
	
#go through and delete the columns that are not in the key parameters			
i = len(keyParameterIndexTupCol) - 1
for z in range( len(parameterNameCol) -1, -1, -1):
	#keyParameterIndexTup[i]
	if (keyParameterIndexTupCol[i] != z):
		ws.delete_rows(z + 1, 1)
	else:
		i = i - 1
		
wb.save("edited.xlsx")