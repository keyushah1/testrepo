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
	
	
''The server name you need to connect to:
'strCurrentEnv = "mysql7003.site4now.net"
' 
''The name of your database:
'dbName = ":db_a359a3_user"
' 
''My SQL connection string. You need to enter your username and password
''strConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver}; Server=" & strCurrentEnv & "; DATABASE="& dbName& ";uid=a359a3_user; pwd=test1234"
''strConnection = "DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=mysql7003.site4now.net; DATABASE=db_a359a3_user;USER=a359a3_user;PASSWORD=test1234;OPTION=3;"
'strConnection = "DRIVER={" &sMySQLDriver& "};SERVER=" &sServerName& "; DATABASE=" &sDatabaseName& ";USER=" &sUserId& ";PASSWORD="&sPassword& ";OPTION=3;"
'Set conn = CreateObject("ADODB.Connection")
'set rs = CreateObject("ADODB.Recordset")
'conn.Open strConnection
' 
''The SQL you want to run
'query = "SELECT * FROM users where username = 'test004'"
' rs.Open query, conn
' 
' If NOT rs.BOF AND NOT rs.EOF Then
' 	MsgBox "Pass"
' Else
'	MsgBox "Fail" 
' End If
' 
''Runs your SQL
''rs = conn.Execute(query)
''MsgBox rs.RecordCount
''If rs.RecordCount > 0 Then
''	MsgBox "Pass"
''Else
''	MsgBox "Fail"
''End If
'dbResults = rs.GetString
'print dbResults
'	
'


'sServerName = "mysql7003.site4now.net"
'sDatabaseName = "db_a359a3_user"
'sMySQLDriver = "MySQL ODBC 5.3 ANSI Driver"
'sUserId = "a359a3_user"
'sPassword = "test1234"
'
'	Set conn = CreateObject("ADODB.Connection")
'	set rs = CreateObject("ADODB.Recordset")
'
'	strConnection = "DRIVER={" &sMySQLDriver& "};SERVER=" &sServerName& "; DATABASE=" &sDatabaseName& ";USER=" &sUserId& ";PASSWORD="&sPassword& ";"
'	
'	'Suppress Errors
'	On Error Resume Next
'		conn.Open strConnection
'	'Turn Errors On	
'	On Error Goto 0
'	
'	If conn.Errors.Count > 0 Then
'		sDBConnectionStatus = "FAIL"
'		verifyRecordEnteredInDB = "FALSE"
'	Else
'		'query = "SELECT * FROM users where username = '" & sUsername & "'"
'		query = "SELECT * FROM users"
'		'query = "delete from users where username like 'test%'"
'	 	rs.Open query, conn
'	 	
'	 	dbResults = rs.GetString
'		print dbResults
'	 
'		 rs.Close
'		 conn.Close
'		 
'		 Set rs = Nothing
'		 Set conn = Nothing
'	End If
