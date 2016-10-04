<%
'Make all variable declarations required
Option Explicit
%>
<!-- #INCLUDE File="ADOVBS.inc" -->
<HTML>
<HEAD>
<TITLE>Beginning SQL Server 2000 for VB Developers</TITLE>
<LINK Rel="StyleSheet" Type="Text/CSS" Href="htPageStyles.css"/>
</HEAD>
<%
'***************************************************************
'* Main process starts here
'***************************************************************

'Declare global variables
Dim intIndex
Dim blnProcessData
Dim objEQCommand, objEQCommands, objEQSession, objEQResponse
Dim objInput, objInputs, objRS
Dim strQuestion, strRestatement, strResponse, strInputName, strInput

'Determine the path of processing
If Request.Form("FormAction") <> "Clarify" Then
	
	'Process normal path, first time into the page
	
	'Save the question from the request form
	strQuestion = Request.Form("txtQuestion")

	'Initialize the English Query session
	Set objEQSession = GetEQSession("Hardware Tracking.eqd")

	'Set the EQ response object to the question
	Set objEQResponse = objEQSession.ParseRequest(strQuestion)

Else
	
	'Process clarification of question
	
	'Set the EQ Response object and Inputs object
	Set objEQResponse = Session("EQResponse")
	Set objInputs = objEQResponse.UserInputs
		
	'Declare constants
	Const eqListInput = 0
	Const eqTextInput = 1
	Const eqStaticInput = 2
		
	'Process all user clarifications
	For intIndex = 0 to objInputs.Count - 1
		'Set the Input object
		Set objInput = objInputs(intIndex)
		'Set the string variables
		strInputName = "Input" & intIndex
		strInput = Request.Form(strInputName)
		'Process the appropriate clarification
		Select Case objInput.Type
			Case eqListInput
				'Input selection chosen
				objInput.Selection = strInput
			Case eqTextInput
				'Input text entered
				objInput.Text = strInput
			Case eqStaticInput
				'Handle static input
			Case Else
				'Handle unknown errors
				Response.Write "Error: Unknown input type"
		End Select
	Next

	'Resubmit the question by replying to the clarification response
	Set objEQResponse = objEQResponse.Reply()

	'Save the question
	strQuestion = Session("Question")

End If

'Process the response
Call ProcessResponse

'###############################################################
'# Main process ends here
'###############################################################

'***************************************************************
'* Subroutines stars here
'***************************************************************

Sub CheckForErrors(objConn)
'***************************************************************
'* Check for and handle ADO errors
'***************************************************************

	'Declare local variables
	Dim blnDisplayErrMsg
	Dim objErr

	If objConn.Errors.Count > 0 Then

		'Create an error object to access the ADO errors collection
		Set objErr = Server.CreateObject("ADODB.Error")

		'Display all errors
		For Each objErr In objConn.Errors
			'Only process errors that are not zero
			If objErr.Number <> 0 Then
				Response.Write objErr.Number & "<br>"
				Response.Write objErr.Description & "<br>"
				Response.Write objErr.Source & "<br>"
				Response.Write objErr.SQLState & "<br>"
				Response.Write objErr.NativeError & "<br><br>"
				blnDisplayErrMsg = True
			End If
		Next

		If blnDisplayErrMsg Then
			'Display a message to the user
			Response.Write "An unforseen error has occurred and " & _
				"processing must be stopped."
			'Halt Execution
			Response.End
		End If

	End IF
End Sub

Function GetEQSession(strDomainFile)
'***************************************************************
'* Get the current English Query Session
'* or 
'* Initialize a new English Query Session
'***************************************************************
	
	If IsObject(Session("EQSession")) Then
		'Get the existing session
		Set objEQSession = Session("EQSession")
	Else
		'Initialize a new session
		'Create the EQ object
		Set objEQSession = Server.CreateObject("MSEQ.Session")
		
		'Initialize the domain
		objEQSession.InitDomain(Server.MapPath(strDomainFile))
		
		'Automatically correct spelling errors when set to False
		
		objEQSession.ClarifySpellingErrors = False
		'Save the EQ session in the Session object for 
		'faster access on subsequent calls
		Set Session("EQSession") = objEQSession
	End If
	
	'Return the EQ session
	Set GetEQSession = objEQSession
End Function

Sub ProcessResponse()
'***************************************************************
'* Process the response from English Query
'***************************************************************
	
	'Declare reponse type constants
	Const eqCommandResponse	= 0
	Const eqErrorResponse = 2
	Const eqUserClarifyResponse = 3

	'Remove reference to existing Session EQ Response object
	Set Session("EQResponse") = Nothing

	'Process EQ response type
	Select Case objEQResponse.Type
		Case eqCommandResponse
			'Normal response
			Call ProcessCommands
			blnProcessData = True
		Case eqErrorResponse
			'Error response
			Response.Write objEQResponse.Description & "<BR>"
			Response.End
		Case eqUserClarifyResponse
			'Clarify response
			blnProcessData = False
		Case Else
			'Unknown response
			Response.Write "Error: Unknown response type<BR>"
			Response.End
	End Select
End Sub

Sub ProcessCommands()
'***************************************************************
' Process all EQ Responses
'***************************************************************

	'Declare command type constants
	Const eqQueryCmd = 1
	Const eqAnswerCmd	= 2

	'Set the strRestatement variable
	strRestatement = "<B>Restatement: </B>" & objEQResponse.Restatement

	'Set the EQ Commands object
	Set objEQCommands = objEQResponse.Commands
	
	'Process the EQ commands
	For intIndex = 0 To objEQCommands.Count - 1
	
		'Set the Command object to the current EQ command
		Set objEQCommand = objEQCommands(intIndex)
		
		'Select and process the appropriate CmdID
		Select Case objEQCommand.CmdID
			Case eqQueryCmd
				'Execute the SQL statement
				Call ProcessSQLCommands
			Case eqAnswerCmd
				'Process the answer
				Response.Write objEQCommand.Answer
			Case Else
				'Process unknown errors
				Response.Write "Error: Unknown command type<BR>"
		End Select
	Next
End Sub

Sub ProcessSQLCommands()
'***************************************************************
'* Process SQL Command
'***************************************************************

	'Ignore errors - we'll handle them
	On Error Resume Next

	'Declare local variables
	Dim intMaxRows
	Dim objConn
	
	'Create the Command object
	Set objConn = Server.CreateObject("ADODB.Connection")
	'Open the connection to SQL Server using a DSN-Less
	'connection w/SQL Server authentication
	objConn.Open "Provider=SQLOLEDB;" & _
				"Data Source=WSTRAVEL\SQL2000;" & _
				"Initial Catalog=Hardware Tracking;" & _
				"User ID=HardwareApplication;" & _
				"Password=hardware;"
	'Check for errors
	Call CheckForErrors(objConn)

	'Set the intMaxRows variable
	intMaxRows = objEQCommand.DisplayRows
	If intMaxRows = 0 Then
		intMaxRows = 25
	End If
	
	'Create the Recordset object
	Set objRS = Server.CreateObject("ADODB.Recordset")
	
	'Use a client side cursor
	objRS.CursorLocation = adUseClient
	
	'Set the maximum number of records to be returned
	objRS.MaxRecords = intMaxRows
	
	'Open the recordset
	objRS.Open objEQCommand.SQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	'Check for errors
	Call CheckForErrors(objConn)
	
	'Disconnect the Recordset object from the Connection object
	Set objRS.ActiveConnection = Nothing
	
	'Close and remove the reference to the Connection object
	objConn.Close
	Set objConn = Nothing
	
	'Save the response from the EQ Command object
	strResponse = "<B>Response: </B>" & objEQCommand.TableCaption
End Sub

Sub RequestClarification()
'***************************************************************
'* Request clarification of question
'***************************************************************

	'Declare local variables
	Dim intItemIndex, intSelection
	Dim strChecked 
	Dim arrItems
	
	'Declare UserInput type constants
	Const eqListInput = 0
	Const eqTextInput = 1
	Const eqStaticInput = 2

	'Ask for clarification and begin the form
	Response.Write "<FORM Action=""Response.asp"" Method=""POST"" " & _ 
		"Name=""frmClarify"">"
	Response.Write "<INPUT Type=""Hidden"" Name=""FormAction"" " & _
		"Value=""Clarify"">"
	Response.Write "<B>Please Clarify:</B> "

	'Set the inputs object to the EQ response user inputs
	Set objInputs = objEQResponse.UserInputs
	
	'Process all user inputs from EQ
	For intIndex = 0 to objInputs.Count - 1
	
		'Set the input object to the user inputs
		Set objInput = objInputs(intIndex)
		
		'Set the input name for the radio button
		strInputName = "Input" & intIndex
		
		'Display the clarification text
		Response.Write objInput.Caption

		Select Case objInput.Type
			Case eqListInput
				'Get an array of items (items are suggested words)
				arrItems = objInput.Items
				
				'Get the default selection
				intSelection = objInput.Selection
				
				'If its less than zero then set it to zero
				If intSelection < 0 Then
					intSelection = 0
				End If
				
				'Build the options of suggested words
				For intItemIndex = 0 To UBound(arrItems, 1)
				
					'Set the strChecked variable to Checked or nothing
					If intItemIndex = intSelection Then
						strChecked = "Checked"
					Else
						strChecked = ""
					End If
					
					'Build the option button
					Response.Write "<BR><INPUT Type=""Radio"" " & _
						"Name=""" & strInputName & """ Value="" " & _
						intItemIndex & """ " & strChecked & ">" & _
						arrItems(intItemIndex)
				Next
				
			Case eqTextInput
				Response.Write "<BR><INPUT Type=""Text"" Name=""" & _
					strInputName & """ Size=""40"">"
					
			Case Else
				'Handle unknown errors
				Response.Write "Error: Unexpected input type"
		End Select
	Next

	'Write two line breaks
	Response.Write "<BR><BR>"
	
	'Build the submit button
	Response.Write "<INPUT Type=""Submit"" Value=""Submit Clarification"">"
	
	'Save the question and response for the clarification page
	Session("Question") = strQuestion
	Set Session("EQResponse") = objEQResponse
End Sub

Sub BuildResultsTable()
'***************************************************************
' Build a table with the results
'***************************************************************

	'Declare local variables
	Dim intCol
	Dim blnEven
	
	'Check for a valid object first
	If IsObject(objRS) Then
	
		'If results exist then process them
		If Not objRS.EOF Then

			Response.Write "<TABLE Border=""0"" Class=""NormalText"">"

			'Build column headers
			Response.Write "<TR>"
			For intCol = 0 To objRS.Fields.Count - 1
				Response.Write "<TD Class=""TableHeader"">" & _
						objRS.Fields.Item(intCol).Name & "</TD>"
			Next
			Response.Write "</TR>"
			
			'Build the rows of data
			blnEven = True
			
			Do While Not objRS.EOF
			
				If blnEven Then
					'Use EvenRow style
					Response.Write "<TR Class=""EvenRow"">"
					'Flip the blnEven variable
					blnEven = Not blnEven
				Else
					'Use OddRow style
					Response.Write "<TR Class=""OddRow"">"
					'Flip the blnEven variable
					blnEven = Not blnEven
				End If
						
				'Process all fields in the Record object
				For intCol = 0 To objRS.Fields.Count - 1
					'Check for a null value
					If IsNull(objRS.Fields(intCol).Value) Then
						'If null then write nothing
						Response.Write "<TD></TD>"
					Else
						'If not null then write the data
						Response.Write "<TD>" & _
							objRS.Fields(intCol).Value & "</TD>"
					End If
				Next

				'Write the closing element for the table row
				Response.Write "</TR>"
				
				'Move to the next record
				objRS.MoveNext
			Loop
			
			'Write the closing element for the table
			Response.Write "</TABLE>"

		Else
			'Display a message that no data was found
			Response.Write("No data was found")
		End If
	
		'Close and remove the reference to the Recordset object
		objRS.Close
		Set objRS = Nothing
		
	End If
End Sub

'###############################################################
'# Subroutines end here
'###############################################################
%>
<BODY>
<TABLE Border="0" Width="100%" Class="NormalText">
	<TR>
		<TH>
			<CENTER>
				Hardware Tracking English Query Application
			</CENTER>
		</TH>
	</TR>
	<TR>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD><B>Question: </B><%=strQuestion%></TD>
	</TR>
	<TR>
		<TD><%=strRestatement%></TD>
	</TR>
	<TR>
		<TD><%=strResponse%></TD>
	</TR>
	<TR>
		<TD>&nbsp;</TD>
	</TR>
	<TR>
		<TD>
<%
	If blnProcessData Then
		'Build the results table
		Call BuildResultsTable
	Else
		'Request clarification
		Call RequestClarification
	End If
%>
		</TD>
	</TR>
	<TR>
		<TD>
			<INPUT Type="Button" Name="btnNewQuestion" 
				Value="New Question">
		</TD>
	</TR>
</TABLE>
<SCRIPT Language="VBScript">
Sub btnNewQuestion_OnClick()
	Window.Location.Href = "EQ.htm"
End Sub
</SCRIPT>
</BODY>
</HTML>
