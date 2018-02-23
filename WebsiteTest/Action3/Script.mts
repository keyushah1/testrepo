'-------------------------------------------------------------------------------------------------------------------
'Action Name: Signup
'Description: This Action tests the Signup feature for all the data rows available in the Signup Data Sheet.
'			  It also performs a database verification to confirm that the user has been registered successfully.
'-------------------------------------------------------------------------------------------------------------------- 


'Get Signup sheet row count
iRowCount = getCountOfDataRows(sExcelPath,"Signup")

'Execute Signup Action for all rows in Signup Sheet
For iRow = 2 To iRowCount
	'Read Signup sheet row content
	readTestData iRow,sExcelPath,"SignUp"
	sUsername=odict.Item("Username")
	sPassword = odict.Item("Password")
	sURL = odict.Item("URL")
	
	'Open New Browser and Navigate to URL	
	Call OpenBrowserAndNavigateURL(sURL)

	If Browser("Sign Up").Page("Sign Up").Exist Then
		Browser("Sign Up").Page("Sign Up").Sync
		If Browser("Sign Up").Page("Sign Up").WebEdit("username").Exist Then
			'Set username
			Browser("Sign Up").Page("Sign Up").WebEdit("username").Set sUsername
			reporterFunc "Done", "Username value set - Username: "&sUsername , ""
			If Browser("Sign Up").Page("Sign Up").WebEdit("password").Exist Then
				'Set Password
				Browser("Sign Up").Page("Sign Up").WebEdit("password").Set sPassword
				reporterFunc "Done", "Password value set - Password: xxxxxxxx" , ""
				If Browser("Sign Up").Page("Sign Up").WebEdit("confirm_password").Exist Then
					'Set Confirm Password
					Browser("Sign Up").Page("Sign Up").WebEdit("confirm_password").Set sPassword
					reporterFunc "Done", "Confirm Password value set - Password: xxxxxxxx" , ""
					If Browser("Sign Up").Page("Sign Up").WebButton("Submit").Exist Then
						'Click Submit
						Browser("Sign Up").Page("Sign Up").WebButton("Submit").Click
						reporterFunc "Done", "Submit Clicked", "Done"
						Wait(5)
						'Verify Login successful
						If Browser("Login").Page("Login").WebElement("Login").Exist Then
							reporterFunc "Pass", "Registration Successful for User: "&sUsername, "Login Page dispayed after successful Registration" 
						Else
							reporterFunc "Fail", "Registration Not Successful", "Login Page NOT dispayed after Registration"
						End If
						'Verify record enetered in database
						If verifyRecordEnteredInDB(sUsername) = "TRUE" Then
							reporterFunc "Pass", "DB Verification for User: "&sUsername, "Registration Details entered in DB" 
						Else
							If sDBConnectionStatus = "FAIL" Then'								
								reporterFunc "Fail", "DB Verification for User: "&sUsername, "DB Connection Failed"
							Else
								reporterFunc "Fail", "DB Verification for User: "&sUsername, "Registration Details NOT available in DB"
							End If
						End If
					Else
						errorLogging "Registration Page", "Submit button does not exist"	
					End If	
				Else
					errorLogging "Registration Page", "Confirm Password field does not exist"	
				End If				
			Else
				errorLogging "Registration Page", "Password field does not exist"	
			End If
		Else
			errorLogging "Registration Page", "Username field does not exist"			
		End If		
	Else
		errorLogging "Registration Page", "Registration Page does not exist"
	End If
	
	SystemUtil.CloseProcessByName "iexplore.exe"
		
Next
