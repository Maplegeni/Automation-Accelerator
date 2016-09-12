' *******************************************************************************
' Name: Class QueryMSSQL
' @Description: 
'1. Validates input parameters
'2. Generates SQL query based on input dictionaries (Results Mapping and Search Criteria)
'3. Contains simplified option to query MS SQL Database - by providing the full SELECT 
'	statement to function GetResultsDictionaryTestQuery
'
'TO USE:

'	1. For the main framework option:
		'Create Search Criteria dictionary:
		
'			Dim dictSearchCriteria: Set dictSearchCriteria = CreateObject("Scripting.Dictionary")
'			dictSearchCriteria.Add "ORDER_NUM", "QZ-000007"

		'Dictionary contains field-value pairs for Search Criteria
		'e.g. this case will be presented as a part of WHERE clause:
		'WHERE ... AND ORDER_NUM = 'QZ-000007'
'
'	2. For the main framework option:
		'Create Filds Mapping dictionary:
		
'			Dim dictResultsMap: Set dictResultsMap = CreateObject("Scripting.Dictionary")
'			dictResultsMap.Add "INCLUDE_COMISN_IND", 	"includeCommInd"
'			dictResultsMap.Add "BOR_ORDER_NUM",			"borOrderNum"
'			dictResultsMap.Add "ESTIMATED_TOTAL_AMT",	"estTotAmt"

		'Dictionary contains field-varName pairs for SELECT clause
		'e.g. this case will be presented as a number of fields of SELECT 
		'SELECT INCLUDE_COMISN_IND, BOR_ORDER_NUM, ESTIMATED_TOTAL_AMT FROM ...
		'The values, returned in dictResult will be available with the "varName part
		'e.g. dictResults.Item("includeCommInd") returns the value of INCLUDE_COMISN_IND field from database (for example "Y" or "N")

'	3. Create an object with QueryMSSQL Factory function (CreateDBQuery()) and setup the mandatory properties:
'			'Create an instanse of QueryMSSQL class
'			Dim oQuery: Set oQuery = CreateDBQuery()
'			'Set properties
'			With oQuery
'				'Mandatory
'				.IP = 			"<DB IP >"
'				.PORT = 		"<DB Port>"
'				.DBUserID = 	"<DB UserName>"
'				.DBPassword = 	"<DB Password>"
'				.DBName = 		"<DB_Name>"
'				.TableName = 	"<DB_Table>"
'				'Optional
'				.Condition = 	"RECORD_TYPE_CD = 'SO'"
'				.Debug = True
'			End With


' Returns:
'	Results Dictiobary object				
' *******************************************************************************

Option Explicit

'Factory
Public Function CreateDBQuery()
	Dim oQuery: Set oQuery = New QueryMSSQL
	Set CreateDBQuery = oQuery
End Function

Class QueryMSSQL
	Private m_sServer_IP
	Private m_sServer_Port
	Private m_dictResults
	Private m_dictSearchCriteria
	Private m_dictResultsMapping
	Private m_sQuery
	Private m_bDebug
	Private m_sDBName
	Private m_sTblName
	Private m_sUserID
	Private m_sPassword
	Private m_oConn
	Private m_oRS
	Private m_iRecordsCount
	Private C_NULL
	Private m_bIsSimpleGet
	Private m_sCondition
	Private m_sDBSchema

	Public Sub Class_Initialize()
		m_sServer_IP = vbNullString
		m_sServer_Port = vbNullString
		C_NULL = vbNullString
		m_sDBName = vbNullString
		m_sDBSchema = vbNullString
		m_sTblName = vbNullString
		m_sUserID = vbNullString   
		m_sPassword = vbNullString 
		Set m_dictResults = CreateObject("Scripting.Dictionary")
		m_sQuery = "SELECT "
		Set m_oConn = Nothing
		Set m_oRS = Nothing
		m_bDebug = False
		m_iRecordsCount = 0
		m_bIsSimpleGet = False
		m_sCondition = vbNullString
	End Sub
	
	Property Let IP(ByVal sIP)
		m_sServer_IP = sIP
	End Property
	
	Property Let PORT(ByVal sPort)
		m_sServer_Port = sPort
	End Property
	
	Property Let DBUserID(ByVal sUserID)
		m_sUserID = sUserID
	End Property
	
	Property Let DBPassword(ByVal sPassword)
		m_sPassword = sPassword
	End Property
	
	Property Let DBName(ByVal sDBName)
		m_sDBName = sDBName
	End Property
	
	Property Let DBSchema(ByVal sDBSchema)
		m_sDBSchema = sDBSchema
	End Property
	
	Property Let TableName(ByVal sTblName)
		m_sTblName = sTblName
	End Property
	
	Property Let Condition(ByVal sCondition)
		m_sCondition = sCondition
	End Property
	
	Property Get RecordCount()
		RecordCount = m_iRecordsCount
	End Property
	
	Property Set SearchCriteria(ByVal dictSearchCriteria)
		Set m_dictSearchCriteria = dictSearchCriteria
	End Property
	
	Property Set ResultsMapping(ByVal dictResultsMapping)
		Set m_dictResultsMapping = dictResultsMapping
	End Property
	
	Property Let Debug(ByVal bIsDebug)
		m_bDebug = bIsDebug
	End Property
	
	'Custom Function for Orders API project
	Public Function GetResultsDictionaryByOrderNum(ByVal sOrderNum)
		Set GetResultsDictionaryByOrderNum = Nothing
		m_sQuery = "SELECT * FROM " & m_sDBSchema & "." & m_sTblName & " WHERE " & m_sCondition & " AND BOR_ORDER_NUM = '" & sOrderNum & "'"
		PrintDebug "Using Customly Generated Query: " & m_sQuery
		If bCreateConn Then
			If bGetRecordset Then
				m_bIsSimpleGet = True
				GenerateResults
				m_bIsSimpleGet = False
				Set GetResultsDictionaryByOrderNum = m_dictResults
			End If
		End If
	End Function
	
	Public Function GetResultsDictionary()
		Set GetResultsDictionary = Nothing
		If bGenerateQuery Then
			If bCreateConn Then
				If bGetRecordset Then
					GenerateResults
					Set GetResultsDictionary = m_dictResults
				End If					
			End If
		End If
	End Function
	
	Public Function GetResultsDictionaryTestQuery(ByVal sQuery)
		PrintDebug "Using Provided Query: " & sQuery
		Set GetResultsDictionaryTestQuery = Nothing
		m_sQuery = sQuery
		If bCreateConn Then
			If bGetRecordset Then
				m_bIsSimpleGet = True
				GenerateResults
				m_bIsSimpleGet = False
				Set GetResultsDictionaryTestQuery = m_dictResults
			End If
		End If
	End Function

	Public Sub PrintResults()
		If IsBlank(m_dictResults) Then
			print "Result Dictionary object is Nothing. No results to print out."
		Else
			PrintDebug "***** Results Dictionary print out *****"
			Dim sKey
			For each sKey in m_dictResults.Keys
				print "Key: " & sKey & "; value: " & m_dictResults(sKey)
			Next
		End If
	End Sub

	Private Function bGenerateQuery()
		bGenerateQuery = False
		If IsBlank(m_dictResultsMapping) Then
			print "Mapping dictionary was not created. Exiting..."
			Exit Function
		ElseIf m_dictResultsMapping.Count = 0 Then
			print "Mapping dictionary must contain at least one key-value pair. Exiting..."
			Exit Function
		End If
		If IsBlank(m_dictSearchCriteria) Then
			print "Search Criteria dictionary was not created. Exiting..."
			Exit Function
		ElseIf m_dictSearchCriteria.Count = 0 Then
			print "Search Criteria dictionary must contain at least one key-value pair. Exiting..."
			Exit Function
		End If
		Dim sQuery: sQuery = vbNullString
		Dim sKey
		Dim bIsFirst: bIsFirst = True
		For each sKey in m_dictResultsMapping.Keys
			If bIsFirst Then
				sQuery = sQuery & sKey
				bIsFirst = False
			Else
				sQuery = sQuery & ", " & sKey
			End If
		Next
		bIsFirst = True
		m_sQuery = m_sQuery & sQuery & " FROM " & m_sDBSchema & "." & m_sTblName & " WHERE " & m_sCondition
		sQuery = vbNullString
		For each sKey in m_dictSearchCriteria.Keys
			If bIsFirst AND IsBlank(m_sCondition) Then
				sQuery = sQuery & sKey & " = '" & m_dictSearchCriteria.Item(sKey) & "'"
				bIsFirst = False
			Else
				sQuery = sQuery & " AND " & sKey & " = '" & m_dictSearchCriteria.Item(sKey) & "'"
			End If
		Next
		m_sQuery = m_sQuery & sQuery
		PrintDebug "Generated Query: " & m_sQuery
		bGenerateQuery = True
	End Function
	
	Private Function bCreateConn()
		bCreateConn = False
		If IsBlank(m_sServer_IP) OR IsBlank(m_sServer_Port) OR IsBlank(m_sDBName) OR IsBlank(m_sUserID) OR IsBlank(m_sPassword) Then
			print "Manatory parameter is not provided. Check Server IP, Port, DB Name, UserID and Password. Exiting..."
			Exit Function
		End If
		Set m_oConn = CreateObject("ADODB.Connection")
		Dim sSQLConn: sSQLConn = "Provider=SQLOLEDB; Data Source=" & m_sServer_IP & "," & m_sServer_Port & "; Initial Catalog=" & m_sDBName & "; Uid=" & m_sUserID & "; Pwd=" & m_sPassword
		PrintDebug "sSQLConn = " & sSQLConn
		m_oConn.ConnectionString = sSQLConn
		On Error Resume Next
		m_oConn.Open
		If err.Number <> 0 Then
			Set m_oConn = Nothing
			print "Unable to create a connection."
			print err.Number & ": " & err.Description
			err.Clear
		Else
			bCreateConn = True
		End If
		On Error GoTo 0
	End Function
	
	Private Function bGetRecordset()
		bGetRecordset = False
		Set m_oRS = CreateObject("ADODB.Recordset")
		m_oRS.CursorLocation = 3 'Using disconnected RS (adUseClient)
		
		On Error Resume Next
		m_oRS.Open m_sQuery, m_oConn
		If err.Number <> 0 Then
			Set m_oRS = Nothing
			print "Unable to get a recordset."
			print err.Number & ": " & err.Description
			err.Clear
		Else
			m_oRS.ActiveConnection = Nothing
			ReleaseConn
			m_oRS.MoveLast
			m_iRecordsCount = m_oRS.RecordCount
			m_oRS.MoveFirst
			bGetRecordset = True
		End If
		On Error GoTo 0
	End Function
	
	Private Sub ReleaseConn()
		m_oConn.Close
		Set m_oConn = Nothing
	End Sub
	
	Private Sub ReleaseRecordset()
		m_oRS.Close
		Set m_oRS = Nothing
	End Sub
	
	Private Sub GenerateResults()
		m_dictResults.RemoveAll
		Dim oField
		Dim sValue: sValue = vbNullString
		If m_iRecordsCount > 0 Then
			If m_bIsSimpleGet Then
				For each oField in m_oRS.Fields
					If IsBlank(oField.value) Then
						m_dictResults.Add oField.Name, C_NULL
					Else
						m_dictResults.Add oField.Name, Trim(CStr(oField.value))
					End If
				Next
			Else
				'While Not m_oRS.EOF
					For each oField in m_oRS.Fields
						If IsBlank(oField.value) Then
							m_dictResults.Add m_dictResultsMapping.Item(oField.Name), C_NULL
						Else
							m_dictResults.Add m_dictResultsMapping.Item(oField.Name), Trim(CStr(oField.value))
						End If
					Next
				'm_oRS.MoveNext	
				'Wend
			End If
			Set oField = Nothing
			ReleaseRecordset
		Else
			print "No records returned."
			Set m_dictResults = Nothing
		End If
		If m_iRecordsCount > 1 Then
			print "Multiple records returned, providing the results for the first record in Recordset"
		End If
	End Sub
	
	Private Function IsBlank(Value)
		'returns True if Empty or NULL or ZeroString
		If IsEmpty(Value) or IsNull(Value) Then
			IsBlank = True
		ElseIf VarType(Value) = vbString Then
			If Value = "" OR Value = vbNullString Then
				IsBlank = True
			End If
		ElseIf IsObject(Value) Then
			If Value Is Nothing Then
				IsBlank = True
			End If
		Else
			IsBlank = False
		End If
	End Function
	
	Private Sub PrintDebug(ByVal sMessage)
		If m_bDebug Then
			print sMessage
		End If
	End Sub
	
End Class

