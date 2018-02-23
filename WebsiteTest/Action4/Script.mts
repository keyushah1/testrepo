'-------------------------------------------------------------------------------------------------------------------
'Action Name: Login
'Description: This Action tests the Login feature for all the data rows available in the Login Data Sheet.
'-------------------------------------------------------------------------------------------------------------------- 


'Get Login sheet row count
iRowCount = getCountOfDataRows(sExcelPath,"Login")
		
'Execute Login Action for all rows in Login Sheet	
For iRow = 2 To iRowCount
	'Read Login sheet row content
	readTestData iRow,sExcelPath,"Login"
	sUsername=odict.Item("Username")
	sPassword = odict.Item("Password")
	sURL = odict.Item("URL")

	'Open New Browser and Navigate to URL
	Call OpenBrowserAndNavigateURL(sURL)

	If Browser("Login").Page("Login").Exist Then
		Browser("Login").Page("Login").Sync
		If Browser("Login").Page("Login").WebEdit("username").Exist Then
			'Set Username
			Browser("Login").Page("Login").WebEdit("username").Set sUsername
			reporterFunc "Done", "Username value set - Username: "&sUsername , ""
			If Browser("Login").Page("Login").WebEdit("password").Exist Then
				'Set Password
				Browser("Login").Page("Login").WebEdit("password").Set sPassword
				reporterFunc "Done", "Password value set - Password: xxxxxxxx" , ""
				If Browser("Login").Page("Login").WebButton("Login").Exist Then
					'Click Login
					Browser("Login").Page("Login").WebButton("Login").Click
					reporterFunc "Done", "Login Clicked", "Done"
					Wait(5)
					'Verify Welcome Text
					If Browser("Welcome").Page("Welcome").WebElement("WelcomeText").Exist Then
						reporterFunc "Pass", "Login Successful for User: "&sUsername, "Welcome Page dispayed after successful Login"
						'Verify Sign Out button
						If Browser("Welcome").Page("Welcome").WebButton("Sign Out of Your Account").Exist Then
							'Click Sign Out
							Browser("Welcome").Page("Welcome").WebButton("Sign Out of Your Account").Click
							Wait(2)
							'Verify Logout Successful
							If Browser("Login").Page("Login").WebEdit("username").Exist Then
								reporterFunc "Pass", "Logout Successful for User: "&sUsername, "Login Page dispayed after successful Logout" 
							Else
								errorLogging "Logout Not Successful for User: "&sUsername, "Login Page NOT displayed asfter successful Logout"			
							End If
						Else
							errorLogging "Welcome Page", "Sign Out button does not exist"
						End If
					Else
						reporterFunc "Fail", "Login Not Successful for User: "&sUsername, "Welcome NOT dispayed after Login"
					End If
				Else
					errorLogging "Login Page", "Login button does not exist"	
				End If
			Else
				errorLogging "Login Page", "Password field does not exist"	
			End If
		Else
			errorLogging "Login Page", "Username field does not exist"			
		End If		
	Else
		errorLogging "Login Page", "Login Page does not exist"
	End If
	
	SystemUtil.CloseProcessByName "iexplore.exe"
		
Next
