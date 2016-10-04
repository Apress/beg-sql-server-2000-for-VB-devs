VERSION 5.00
Begin VB.Form frmComparison 
   Caption         =   "Side-by-Side Comparisons"
   ClientHeight    =   3816
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3816
   ScaleWidth      =   4104
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSQL 
      Height          =   1068
      Left            =   96
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2688
      Width           =   3948
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview SQL"
      Height          =   396
      Left            =   2976
      TabIndex        =   7
      Top             =   96
      Width           =   1068
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   396
      Left            =   2976
      TabIndex        =   8
      Top             =   576
      Width           =   1068
   End
   Begin VB.Frame Frame1 
      Caption         =   "Execution Mode"
      Height          =   2028
      Left            =   96
      TabIndex        =   11
      Top             =   576
      Width           =   2508
      Begin VB.OptionButton optExecution 
         Caption         =   "Update In-Line SQL"
         Height          =   204
         Index           =   5
         Left            =   192
         TabIndex        =   6
         Top             =   1728
         Width           =   2004
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Update Stored Procedure"
         Height          =   204
         Index           =   4
         Left            =   192
         TabIndex        =   5
         Top             =   1440
         Width           =   2100
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Insert In-Line SQL"
         Height          =   204
         Index           =   3
         Left            =   192
         TabIndex        =   4
         Top             =   1152
         Width           =   2004
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Insert Stored Procedure"
         Height          =   204
         Index           =   2
         Left            =   192
         TabIndex        =   3
         Top             =   864
         Width           =   2004
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Select In-Line SQL"
         Height          =   204
         Index           =   1
         Left            =   192
         TabIndex        =   2
         Top             =   576
         Width           =   2004
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Select Stored Procedure"
         Height          =   204
         Index           =   0
         Left            =   192
         TabIndex        =   1
         Top             =   288
         Value           =   -1  'True
         Width           =   2004
      End
   End
   Begin VB.ComboBox cboCDTypes 
      Height          =   288
      Left            =   960
      TabIndex        =   0
      Top             =   156
      Width           =   1644
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      CausesValidation=   0   'False
      Height          =   396
      Left            =   2976
      TabIndex        =   9
      Top             =   1056
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "CD Types"
      Height          =   204
      Left            =   96
      TabIndex        =   10
      Top             =   192
      Width           =   972
   End
End
Attribute VB_Name = "frmComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Delcare variables
Dim intExecutionMode As Integer
Dim lngCDID As Long

Sub ADOError()
   'Declare local variables
   Dim strError As String
   
   'Loop through the errors collection and display all errors
   For Each g_objError In g_objConn.Errors
      strError = strError & g_objError.Number & " : " & _
         g_objError.Description & vbCrLf & vbCrLf
   Next
   MsgBox strError, vbCritical + vbOKOnly, "ADO Error"
End Sub


Sub InsertSQL()
   'Validate CD type data
   If Len(Trim(cboCDTypes.Text)) = 0 Then
      MsgBox "You must enter a CD type to add.", _
         vbInformation + vbOKOnly, "Insert SQL"
      cboCDTypes.SetFocus
      Exit Sub
   End If

   'Setup error handling
   On Error GoTo InsertSQL_Err
   
   'Declare and set a reference to the Command object
   Dim objCmd As New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = g_objConn
   objCmd.CommandText = "INSERT INTO CD_T " & _
      "(CD_Type_CH, Last_Update_DT) " & _
      "VALUES('" & cboCDTypes.Text & "', '" & Now & "')"
   objCmd.CommandType = adCmdText
   
   'Execute the command object to insert the data
   objCmd.Execute
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset
   g_objRS.Open "SELECT MAX(CD_ID) AS 'CD_ID' FROM CD_T", _
      g_objConn, adOpenForwardOnly, adLockReadOnly, adCmdText

   'Display the Identity value that was inserted for this entry
   txtSQL.Text = "The Identity value that was inserted using" & vbCrLf & _
      "in-line SQL statements is " & g_objRS!CD_ID
   
   'Close and derefernce the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
   
   'Remove the reference to the command object
   Set objCmd = Nothing
   
   Exit Sub
   
InsertSQL_Err:
   'Call the error routine
   Call ADOError
End Sub

Sub InsertStoredProcedure()
   'Validate CD type data
   If Len(Trim(cboCDTypes.Text)) = 0 Then
      MsgBox "You must enter a CD type to add.", _
         vbInformation + vbOKOnly, "Insert Stored Procedure"
      cboCDTypes.SetFocus
      Exit Sub
   End If

   'Setup error handling
   On Error GoTo InsertStoredProcedure_Err
   
   'Declare and set a reference to the Command object
   Dim objCmd As New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = g_objConn
   objCmd.CommandText = "up_parmins_cd_type"
   objCmd.CommandType = adCmdStoredProc
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("CD Type", _
      adChar, adParamInput, 4, cboCDTypes.Text)
   
   'Execute the command object to insert the data
   objCmd.Execute
   
   'Display the Identity value that was inserted for this entry
   txtSQL.Text = "The Identity value that was inserted using" & vbCrLf & _
      "the up_parmins_cd_type stored procedure" & vbCrLf & _
      "is " & objCmd.Parameters.Item("RC")
   
   'Remove the reference to the command object
   Set objCmd = Nothing
   
   Exit Sub
   
InsertStoredProcedure_Err:
   'Call the error routine
   Call ADOError
End Sub

Sub SelectSQL()
   'Setup error handling
   On Error GoTo SelectSQL_Err
   
   'Declare local variables
   Dim strSQL As String
   
   'Build SQL string
   strSQL = "SELECT CD_ID, CD_Type_CH " & _
      "FROM CD_T " & _
      "ORDER BY CD_Type_CH"
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset object
   g_objRS.Open strSQL, g_objConn, adOpenForwardOnly, _
      adLockReadOnly, adCmdText
   
   'Clear the combo box of any previous entries
   cboCDTypes.Clear
   
   'Loop through the recordset and load the cd combo box
   Do While Not g_objRS.EOF
      'Add the cd types
      cboCDTypes.AddItem g_objRS!CD_Type_CH
      'Add the cd id
      cboCDTypes.ItemData(cboCDTypes.NewIndex) = g_objRS!CD_ID
      'Move to the next record
      g_objRS.MoveNext
   Loop
   
   'Close and dereference the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
   
   Exit Sub
   
SelectSQL_Err:
   'Call the error routine
   Call ADOError
End Sub

Sub UpdateSQL()
   'Validate CD type data
   If Len(Trim(cboCDTypes.Text)) = 0 Then
      MsgBox "You must select and change a CD type to update.", _
         vbInformation + vbOKOnly, "Update SQL"
      cboCDTypes.SetFocus
      Exit Sub
   End If

   'Setup error handling
   On Error GoTo UpdateSQL_Err
   
   'Declare local variables
   Dim strSQL As String
   
   'Build the SQL string
   strSQL = "UPDATE CD_T " & _
      "SET CD_Type_CH = '" & RTrim(cboCDTypes.Text) & "', " & _
      "Last_Update_DT = '" & Now & "' " & _
      "WHERE CD_ID = " & lngCDID
      
   'Execute the SQL string
   g_objConn.Execute strSQL

   'Display the Identity value that was inserted for this entry
   txtSQL.Text = "Update Complete"
   
   Exit Sub
   
UpdateSQL_Err:
   'Call the error routine
   Call ADOError
End Sub

Sub UpdateStoredProcedure()
   'Validate CD type data
   If Len(Trim(cboCDTypes.Text)) = 0 Then
      MsgBox "You must select and change a CD type to update.", _
         vbInformation + vbOKOnly, "Update Stored Procedure"
      cboCDTypes.SetFocus
      Exit Sub
   End If

   'Setup error handling
   On Error GoTo UpdateStoredProcedure_Err
   
   'Declare and set a reference to the Command object
   Dim objCmd As New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = g_objConn
   objCmd.CommandText = "up_parmupd_cd_type"
   objCmd.CommandType = adCmdStoredProc
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("CD ID", _
      adInteger, adParamInput, , lngCDID)
   objCmd.Parameters.Append objCmd.CreateParameter("CD Type", _
      adChar, adParamInput, 4, RTrim(cboCDTypes.Text))
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamOutput)
   
   'Execute the command object to insert the data
   objCmd.Execute
   
   'Display the return code
   txtSQL.Text = "The return code from the up_parmupd_cd_type" & vbCrLf & _
      "stored procedure is " & objCmd("RC")
   
   'Remove the reference to the command object
   Set objCmd = Nothing
   
   Exit Sub
   
UpdateStoredProcedure_Err:
   'Call the error routine
   Call ADOError
End Sub

Private Sub cboCDTypes_Click()
   'Save the ID of the CD type in case we need it
   If cboCDTypes.ListIndex > -1 Then
      lngCDID = cboCDTypes.ItemData(cboCDTypes.ListIndex)
   End If
End Sub

Private Sub cmdExecute_Click()
   Select Case intExecutionMode
      Case 0
         'Execute Select Stored Procedure
         Call SelectStoredProcedure
      Case 1
         'Execute Select SQL String
         Call SelectSQL
      Case 2
         'Execute Insert Stored Procedure
         Call InsertStoredProcedure
      Case 3
         'Execute Insert SQL String
         Call InsertSQL
      Case 4
         'Execute Update Stored Procedure
         Call UpdateStoredProcedure
      Case 5
         'Execute Update SQL String
         Call UpdateSQL
   End Select
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub SelectStoredProcedure()
   'Setup error handling
   On Error GoTo SelectStoredProcedure_Err
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset object
   g_objRS.Open "up_select_cd_types", g_objConn, adOpenForwardOnly, _
      adLockReadOnly, adCmdStoredProc
   
   'Clear the combo box of any previous entries
   cboCDTypes.Clear
   
   'Loop through the recordset and load the cd combo box
   Do While Not g_objRS.EOF
      'Add the cd types
      cboCDTypes.AddItem g_objRS!CD_Type_CH
      'Add the cd id
      cboCDTypes.ItemData(cboCDTypes.NewIndex) = g_objRS!CD_ID
      'Move to the next record
      g_objRS.MoveNext
   Loop
   
   'Close and dereference the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
   
   Exit Sub
   
SelectStoredProcedure_Err:
   'Call the error routine
   Call ADOError
End Sub



Private Sub cmdPreview_Click()
   Select Case intExecutionMode
      Case 0
         'Display Select Stored Procedure
         txtSQL.Text = "up_select_cd_types"
      Case 1
         'Display Select SQL String
         txtSQL.Text = "SELECT CD_ID, CD_Type_CH " & vbCrLf & _
            "FROM CD_T " & vbCrLf & _
            "ORDER BY CD_Type_CH"
      Case 2
         'Display Insert Stored Procedure
         txtSQL.Text = "up_parmins_cd_type"
      Case 3
         'Display Insert SQL String
         txtSQL.Text = "INSERT INTO CD_T " & vbCrLf & _
            "(CD_Type_CH, Last_Update_DT) " & vbCrLf & _
            "VALUES('" & cboCDTypes.Text & "', '" & Now & "')" & vbCrLf & vbCrLf & _
            "SELECT MAX(CD_ID) AS 'CD_ID' FROM CD_T"
      Case 4
         'Display Update Stored Procedure
         txtSQL.Text = "up_parmupd_cd_type"
      Case 5
         'Display Update SQL String
         txtSQL.Text = "UPDATE CD_T " & vbCrLf & _
            "SET CD_Type_CH = '" & RTrim(cboCDTypes.Text) & "', " & vbCrLf & _
            "Last_Update_DT = '" & Now & "' " & vbCrLf & _
            "WHERE CD_ID = " & lngCDID
   End Select
End Sub

Private Sub Form_Load()
   'Display the login form
   frmLogin.Show vbModal
   
   'Set default execution mode
   intExecutionMode = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Termination the database connection
   Call TerminateConnection
End Sub


Private Sub optExecution_Click(Index As Integer)
   'Save the option chosen
   intExecutionMode = Index
End Sub


