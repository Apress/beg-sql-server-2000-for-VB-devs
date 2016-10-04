VERSION 5.00
Begin VB.Form frmSpeedTest 
   Caption         =   "Speed Test"
   ClientHeight    =   1056
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4752
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1056
   ScaleWidth      =   4752
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   396
      Left            =   3744
      TabIndex        =   1
      Top             =   576
      Width           =   972
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   396
      Left            =   3744
      TabIndex        =   0
      Top             =   96
      Width           =   972
   End
   Begin VB.Label lblPrepared 
      Caption         =   "Prepared"
      Height          =   204
      Left            =   96
      TabIndex        =   3
      Top             =   480
      Width           =   3564
   End
   Begin VB.Label lblNonPrepared 
      Caption         =   "Non Prepared"
      Height          =   204
      Left            =   96
      TabIndex        =   2
      Top             =   192
      Width           =   3468
   End
End
Attribute VB_Name = "frmSpeedTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdRun_Click()
   'Declare command object for unprepared execution
   Dim objCmd1 As New ADODB.Command
   
   'Declare command object for prepared execution
   Dim objCmd2 As New ADODB.Command
   
   'Declare variables to hold the start and end times
   Dim sngStart As Single
   Dim sngEnd As Single
   
   'Declare variable for loop counter
   Dim intLoop As Integer
   
   'Set the first command objects properties
   Set objCmd1.ActiveConnection = g_objConn
   objCmd1.CommandText = "SELECT Software_Category_ID, " & _
      "Software_Category_VC " & _
      "FROM Software_Category_T"
   objCmd1.CommandType = adCmdText
   
   'Save the start time
   sngStart = Timer
   
   'Execute the loop
   For intLoop = 1 To 1000
      objCmd1.Execute
   Next intLoop
   
   'Save the end time
   sngEnd = Timer
   
   'Display the results
   lblNonPrepared.Caption = "Non prepared execution ran for " & _
      Format(sngEnd - sngStart, "#0.000") & " seconds"
   
   'Set the second command object's properties
   Set objCmd2.ActiveConnection = g_objConn
   objCmd2.CommandText = "SELECT Software_ID, Software_Name_VC " & _
      "FROM Software_T"
   objCmd2.CommandType = adCmdText
   objCmd2.Prepared = True
   
   'Save the start time
   sngStart = Timer
   
   'Execute the loop
   For intLoop = 1 To 1000
      objCmd2.Execute
   Next intLoop
   
   'Save the end time
   sngEnd = Timer

   'Display the results
   lblPrepared.Caption = "Prepared execution ran for " & _
      Format(sngEnd - sngStart, "#0.000") & " seconds"
      
   'Dereference command objects
   Set objCmd1 = Nothing
   Set objCmd2 = Nothing
End Sub


Private Sub Form_Load()
   'Display the login form
   frmLogin.Show vbModal
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Termination the database connection
   Call TerminateConnection
End Sub


