VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3960
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5088
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5088
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   396
      Left            =   3840
      TabIndex        =   13
      Top             =   3456
      Width           =   1164
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   396
      Left            =   2496
      TabIndex        =   12
      Top             =   3456
      Width           =   1164
   End
   Begin VB.Frame Frame2 
      Caption         =   "Authentication Mode"
      Height          =   2124
      Left            =   96
      TabIndex        =   4
      Top             =   1248
      Width           =   4908
      Begin VB.CheckBox chkConnectString 
         Caption         =   "Make the UserID and Password Part of the Connect String"
         Enabled         =   0   'False
         Height          =   300
         Left            =   480
         TabIndex        =   11
         Top             =   1728
         Width           =   4332
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   1632
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1302
         Width           =   1452
      End
      Begin VB.TextBox txtLogin 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   288
         Left            =   1632
         TabIndex        =   8
         Top             =   918
         Width           =   1452
      End
      Begin VB.OptionButton optMode 
         Caption         =   "SQL Server Authentication"
         Height          =   300
         Index           =   1
         Left            =   192
         TabIndex        =   6
         Top             =   576
         Width           =   2220
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Windows Authentication"
         Height          =   300
         Index           =   0
         Left            =   192
         TabIndex        =   5
         Top             =   288
         Value           =   -1  'True
         Width           =   2220
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Enabled         =   0   'False
         Height          =   204
         Left            =   480
         TabIndex        =   9
         Top             =   1344
         Width           =   876
      End
      Begin VB.Label lblLogin 
         Caption         =   "Login Name"
         Enabled         =   0   'False
         Height          =   204
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   972
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DSN"
      Height          =   1068
      Left            =   96
      TabIndex        =   0
      Top             =   96
      Width           =   4908
      Begin VB.TextBox txtDSN 
         Height          =   288
         Left            =   1152
         TabIndex        =   3
         Text            =   "HardwareAppl"
         Top             =   246
         Width           =   1356
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "DSN-Less"
         Height          =   204
         Index           =   1
         Left            =   192
         TabIndex        =   2
         Top             =   576
         Width           =   1068
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "DSN"
         Height          =   204
         Index           =   0
         Left            =   192
         TabIndex        =   1
         Top             =   288
         Value           =   -1  'True
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   'Establish connection to SQL Server
   If EstablishConnection(blnDSN:=optDSN(0).Value, _
      strDSN:=txtDSN.Text, _
      blnWindowsAuthentication:=optMode(0).Value, _
      strLogin:=txtLogin.Text, _
      strPassword:=txtPassword.Text, _
      blnConnectString:=chkConnectString.Value) Then
      MsgBox "Login Successful!", vbInformation + vbOKOnly, "Login"
   Else
      MsgBox "Connection Failed", vbCritical + vbOKOnly, "Login Failed"
   End If
   
   'Close connection
   Call TerminateConnection
End Sub

Private Sub optDSN_Click(Index As Integer)
   If Index = 0 Then
      'If DSN then enable fields
      txtDSN.Enabled = True
      'Use ColorConstants enumeration
      txtDSN.BackColor = vbWhite
   Else
      'If DSN-Less then disable fields
      txtDSN.Enabled = False
      'Use SystemColorConstants enumeration
      txtDSN.BackColor = vbInactiveCaptionText
   End If
End Sub


Private Sub optMode_Click(Index As Integer)
   If Index = 0 Then
      'If Windows Authentication then disable fields
      lblLogin.Enabled = False
      txtLogin.Enabled = False
      'Use SystemColorConstants enumeration
      txtLogin.BackColor = vbInactiveCaptionText
      lblPassword.Enabled = False
      txtPassword.Enabled = False
      'Use SystemColorConstants enumeration
      txtPassword.BackColor = vbInactiveCaptionText
      chkConnectString.Enabled = False
   Else
      'If SQL Server Authentication then enable fields
      lblLogin.Enabled = True
      txtLogin.Enabled = True
      'Use ColorConstants enumeration
      txtLogin.BackColor = vbWhite
      lblPassword.Enabled = True
      txtPassword.Enabled = True
      'Use ColorConstants enumeration
      txtPassword.BackColor = vbWhite
      chkConnectString.Enabled = True
   End If
End Sub


