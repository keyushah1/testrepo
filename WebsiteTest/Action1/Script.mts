'-------------------------------------------------------------------------------------------------------------------
'Action Name: Master Driver
'Description: This Action is the driver for all the tests. It reads the tests to be run from the TestData.xls sheet. 
'--------------------------------------------------------------------------------------------------------------------             


'Get MasterDataSheet row count
iMasterRowCount = getCountOfMasterDataRows(sExcelPath,sMasterDataSheetName)

'Loop though all rows in MasterDataSheet and execute tests
For iMasterRow = 2 To iMasterRowCount
	'Read MasterData row content
	readMasterData iMasterRow,sExcelPath,sMasterDataSheetName	
	sExecutionFlag=oMasterDict.Item("Execution Flag")
	sActionName = oMasterDict.Item("Action Name")
	'Call Action if Execution Flag is set to Yes
	If sExecutionFlag = "Yes" Then	
		RunAction sActionName, oneIteration
	End If
Next





