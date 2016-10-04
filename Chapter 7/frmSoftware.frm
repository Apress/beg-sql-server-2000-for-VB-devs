VERSION 5.00
Begin VB.Form frmSoftware 
   Caption         =   "Software"
   ClientHeight    =   1920
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4008
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4008
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   396
      Left            =   1248
      TabIndex        =   8
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   396
      Left            =   1248
      TabIndex        =   7
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   396
      Left            =   96
      TabIndex        =   6
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   396
      Left            =   2976
      TabIndex        =   5
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   396
      Left            =   96
      TabIndex        =   4
      Top             =   960
      Width           =   972
   End
   Begin VB.ComboBox cboCategory 
      Height          =   288
      Left            =   1536
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2412
   End
   Begin VB.ComboBox cboSoftware 
      Height          =   288
      Left            =   1536
      TabIndex        =   1
      Top             =   150
      Width           =   2412
   End
   Begin VB.Label Label2 
      Caption         =   "Software Category"
      Height          =   204
      Left            =   96
      TabIndex        =   2
      Top             =   522
      Width           =   1356
   End
   Begin VB.Label Label1 
      Caption         =   "Software Title"
      Height          =   204
      Left            =   96
      TabIndex        =   0
      Top             =   192
      Width           =   972
   End
End
Attribute VB_Name = "frmSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare variables
Dim arrSoftware() As Integer
Dim lngSoftwareID As Long

Private Sub cboSoftware_Click()
   'Declare local variables
   Dim intIndex As Integer
   
   'Loop through the category combo box until a match is found
   For intIndex = 0 To cboCategory.ListCount - 1
      If cboCategory.ItemData(intIndex) = arrSoftware(cboSoftware.ListIndex) Then
         cboCategory.ListIndex = intIndex
         Exit For
      End If
   Next
   
   'Save the key
   lngSoftwareID = cboSoftware.ItemData(cboSoftware.ListIndex)
End Sub


Private Sub cmdDelete_Click()
   'Validate software title was selected
   If cboSoftware.ListIndex = -1 Then
      MsgBox "You must select the software title to delete.", _
         vbInformation + vbOKOnly, "Software Delete"
      cboSoftware.SetFocus
      Exit Sub
   End If

   'Declare local variables
   Dim strSQL As String
   
   'Build the SQL string
   strSQL = "up_parmdel_software " & cboSoftware.ItemData(cboSoftware.ListIndex)
   
   'Execute the SQL string
   g_objConn.Execute strSQL
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdInsert_Click()
   'Validate software title was entered
   If Len(Trim(cboSoftware.Text)) = 0 Then
      MsgBox "You must enter a software title to add.", _
         vbInformation + vbOKOnly, "Software Insert"
      cboSoftware.SetFocus
      Exit Sub
   End If
   
   'Validate software cateogry was selected
   If cboCategory.ListIndex = -1 Then
      MsgBox "You must select the appropriate software category.", _
         vbInformation + vbOKOnly, "Software Insert"
      cboCategory.SetFocus
      Exit Sub
   End If
   
   'Declare and set a reference to the Command object
   Dim objCmd As New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = g_objConn
   objCmd.CommandText = "up_parmins_software"
   objCmd.CommandType = adCmdStoredProc
   
   'Append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("Software", _
      adVarChar, adParamInput, 30, cboSoftware.Text)
   objCmd.Parameters.Append objCmd.CreateParameter("Category", _
      adInteger, adParamInput, , cboCategory.ItemData(cboCategory.ListIndex))
   
   'Execute the command object to insert the data
   objCmd.Execute
   
   'Remove the reference to the command object
   Set objCmd = Nothing
End Sub

Private Sub cmdSelect_Click()
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset object
   g_objRS.Open "up_select_software", g_objConn, adOpenStatic, _
      adLockReadOnly, adCmdStoredProc
      
   'Move to the last record so we can get an accurate count of the
   'number of records in the recordset object
   g_objRS.MoveLast
   
   'Redim the software array to the correct number of entries
   ReDim arrSoftware(g_objRS.RecordCount - 1)
   
   'Now move to the first record to prepare for reading
   g_objRS.MoveFirst
   
   'Clear any existing entries
   cboSoftware.Clear
   
   'Loop through the recordset and load the software combo box
   Do While Not g_objRS.EOF
      'Add the software name
      cboSoftware.AddItem g_objRS!software_name_vc
      'Add the software id
      cboSoftware.ItemData(cboSoftware.NewIndex) = g_objRS!software_id
      'Add the software category id to the software array
      arrSoftware(cboSoftware.NewIndex) = g_objRS!software_category_id
      'Move to the next record
      g_objRS.MoveNext
   Loop
   
   'Close and dereference the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
End Sub


Private Sub cmdUpdate_Click()
   'Validate software title was entered
   If Len(Trim(cboSoftware.Text)) = 0 Then
      MsgBox "You must enter the new software title.", _
         vbInformation + vbOKOnly, "Software Update"
      cboSoftware.SetFocus
      Exit Sub
   End If
   
   'Validate software cateogry was selected
   If cboCategory.ListIndex = -1 Then
      MsgBox "You must select the appropriate software category.", _
         vbInformation + vbOKOnly, "Software Update"
      cboCategory.SetFocus
      Exit Sub
   End If
   
   'Declare and set a reference to the Command object
   Dim objCmd As New ADODB.Command
   
   'Declare the Parameter object
   Dim objParm As ADODB.Parameter
   
   'Set the command object properties
   Set objCmd.ActiveConnection = g_objConn
   objCmd.CommandText = "up_parmupd_software"
   objCmd.CommandType = adCmdStoredProc
   
   'Set a reference to the ADO Parameter object and then set
   'the Parameter's properties
   Set objParm = New ADODB.Parameter
   objParm.Name = "Software ID"
   objParm.Type = adInteger
   objParm.Direction = adParamInput
   objParm.Value = lngSoftwareID
   
   'Append the Parameter to the Parameters collection
   objCmd.Parameters.Append objParm
   
   'Set a reference to the ADO Parameter object and then set
   'the Parameter's properties
   Set objParm = New ADODB.Parameter
   objParm.Name = "Software Title"
   objParm.Type = adVarChar
   objParm.Direction = adParamInput
   objParm.Size = 30
   objParm.Value = cboSoftware.Text
   
   'Append the Parameter to the Parameters collection
   objCmd.Parameters.Append objParm
  
   'Set a reference to the ADO Parameter object and then set
   'the Parameter's properties
   Set objParm = New ADODB.Parameter
   objParm.Name = "Category ID"
   objParm.Type = adInteger
   objParm.Direction = adParamInput
   objParm.Value = cboCategory.ItemData(cboCategory.ListIndex)
   
   'Append the Parameter to the Parameters collection
   objCmd.Parameters.Append objParm
   
   'Execute the command object to update the data
   objCmd.Execute
   
   'Remove the references to the ADO objects
   Set objParm = Nothing
   Set objCmd = Nothing
End Sub

Private Sub Form_Load()
   'Display the login form
   frmLogin.Show vbModal
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset object
   g_objRS.Open "up_select_categories", g_objConn, adOpenForwardOnly, _
      adLockReadOnly, adCmdStoredProc
   
   'Loop through the recordset and load the category combo box
   Do While Not g_objRS.EOF
      'Add the category name
      cboCategory.AddItem g_objRS!software_category_vc
      'Add the category id
      cboCategory.ItemData(cboCategory.NewIndex) = g_objRS!software_category_id
      'Move to the next record
      g_objRS.MoveNext
   Loop
   
   'Close and dereference the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Terminate database connection
   Call TerminateConnection
End Sub


