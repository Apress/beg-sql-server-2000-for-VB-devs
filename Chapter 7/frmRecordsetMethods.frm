VERSION 5.00
Begin VB.Form frmRecordsetMethods 
   Caption         =   "Recordset Methods"
   ClientHeight    =   2808
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6912
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2808
   ScaleWidth      =   6912
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearLists 
      Caption         =   "Clear Lists"
      Height          =   396
      Left            =   4512
      TabIndex        =   8
      Top             =   384
      Width           =   2316
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   396
      Left            =   4512
      TabIndex        =   7
      Top             =   2304
      Width           =   2316
   End
   Begin VB.CommandButton cmdOpenSavedRecordset 
      Caption         =   "Open Saved Recordsets"
      Height          =   396
      Left            =   4512
      TabIndex        =   6
      Top             =   1824
      Width           =   2316
   End
   Begin VB.CommandButton cmdSaveRecordset 
      Caption         =   "Save Recordsets"
      Height          =   396
      Left            =   4512
      TabIndex        =   5
      Top             =   1344
      Width           =   2316
   End
   Begin VB.CommandButton cmdOpenDoubleRecordset 
      Caption         =   "Open Double Recordset"
      Height          =   396
      Left            =   4512
      TabIndex        =   4
      Top             =   864
      Width           =   2316
   End
   Begin VB.ListBox lstCategories 
      Height          =   2352
      Left            =   2592
      TabIndex        =   3
      Top             =   384
      Width           =   1740
   End
   Begin VB.ListBox lstSoftware 
      Height          =   2352
      Left            =   96
      TabIndex        =   0
      Top             =   384
      Width           =   2412
   End
   Begin VB.Label Label2 
      Caption         =   "Categories"
      Height          =   204
      Left            =   2592
      TabIndex        =   2
      Top             =   96
      Width           =   876
   End
   Begin VB.Label Label1 
      Caption         =   "Software"
      Height          =   204
      Left            =   96
      TabIndex        =   1
      Top             =   96
      Width           =   780
   End
End
Attribute VB_Name = "frmRecordsetMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CloseRecordset()
   'Close and dereference the recordset object
   g_objRS.Close
   Set g_objRS = Nothing
End Sub


Sub GetSoftwareAndCategories()
   'Setup error handling
   On Error GoTo GetSoftwareAndCategories_Err
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the recordset object
   g_objRS.Open "up_select_software_and_categories", g_objConn, _
      adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

   Exit Sub
   
GetSoftwareAndCategories_Err:
   'Call the error routine
   Call ErrorHandler
End Sub


Sub LoadCategories()
   'Loop through the recordset and load the categories list box
   Do While Not g_objRS.EOF
      'Add the software category
      lstCategories.AddItem g_objRS!Software_Category_VC
      'Move to the next record
      g_objRS.MoveNext
   Loop
End Sub

Sub LoadSoftware()
   'Loop through the recordset and load the software list box
   Do While Not g_objRS.EOF
      'Add the software title
      lstSoftware.AddItem g_objRS!Software_Name_VC
      'Move to the next record
      g_objRS.MoveNext
   Loop
End Sub

Private Sub cmdClearLists_Click()
   'Clear the list boxes
   lstSoftware.Clear
   lstCategories.Clear
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOpenDoubleRecordset_Click()
   'Setup error handling
   On Error GoTo cmdOpenDoubleRecordset_Err
   
   'Open the recordset
   Call GetSoftwareAndCategories
   
   'Load the software list
   Call LoadSoftware
   
   'Get the next recordset
   Set g_objRS = g_objRS.NextRecordset
   
   'Load the categories list
   Call LoadCategories
   
   'Close and dereference the recordset object
   Call CloseRecordset
   
   Exit Sub
   
cmdOpenDoubleRecordset_Err:
   'Call the error routine
   Call ErrorHandler
End Sub

Sub ErrorHandler()
   'Declare local variables
   Dim strError
   
   'Loop through the errors collection and display all errors
   For Each g_objError In g_objConn.Errors
      strError = strError & g_objError.Number & " : " & _
         g_objError.Description & vbCrLf & vbCrLf
   Next
   
   'Ensure this is an ADO error
   If IsEmpty(strError) Then
      MsgBox "Error " & Err.Number & vbCrLf & vbCrLf & Err.Description, _
         vbCritical + vbOKOnly, "VB Error"
   Else
      MsgBox strError, vbCritical + vbOKOnly, "ADO Error"
   End If
End Sub

Private Sub cmdOpenSavedRecordset_Click()
   'Setup error handling
   On Error GoTo cmdSaveRecordset_Err
   
   'Set a reference to the ADO recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Open the saved software recordset
   g_objRS.Open "Software.rs"

   'Load the software list
   Call LoadSoftware
   
   'Close the current recordset
   g_objRS.Close
   
   'Open the saved categories recordset
   g_objRS.Open "Categories.rs"
   
   'Load the categories list
   Call LoadCategories
   
   'Close and dereference the recordset object
   Call CloseRecordset
   
   Exit Sub
   
cmdSaveRecordset_Err:
   'Call the error routine
   Call ErrorHandler
End Sub

Private Sub cmdSaveRecordset_Click()
   'Setup error handling
   On Error GoTo cmdSaveRecordset_Err
   
   'Open the recordset
   Call GetSoftwareAndCategories
   
   'Save the software recordset
   g_objRS.Save "Software.rs", adPersistADTG

   'Get the next recordset
   Set g_objRS = g_objRS.NextRecordset
   
   'Save the categories recordset
   g_objRS.Save "Categories.rs", adPersistADTG

   'Close and dereference the recordset object
   Call CloseRecordset
   
   Exit Sub
   
cmdSaveRecordset_Err:
   'Call the error routine
   Call ErrorHandler
End Sub

Private Sub Form_Load()
   'Display the login form
   frmLogin.Show vbModal
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Termination the database connection
   Call TerminateConnection
End Sub




