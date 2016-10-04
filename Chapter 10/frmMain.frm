VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmMain 
   Caption         =   "Hardware Tracking"
   ClientHeight    =   5376
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   8328
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5376
   ScaleWidth      =   8328
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   396
      Left            =   4800
      TabIndex        =   52
      Top             =   4512
      Width           =   972
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   396
      Left            =   3648
      TabIndex        =   51
      Top             =   4512
      Width           =   972
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   396
      Left            =   2496
      TabIndex        =   50
      Top             =   4512
      Width           =   972
   End
   Begin TabDlg.SSTab tabData 
      Height          =   3852
      Left            =   96
      TabIndex        =   1
      Top             =   480
      Width           =   8148
      _ExtentX        =   14372
      _ExtentY        =   6795
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Employee"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Hardware"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Software"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "System Assignment"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "System Assignment"
         Height          =   3276
         Left            =   -74808
         TabIndex        =   41
         Top             =   384
         Width           =   7788
         Begin VB.ListBox lstInstalledSoftware 
            Height          =   2640
            Left            =   4320
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   444
            Width           =   3276
         End
         Begin VB.TextBox txtSystemSerialNumber 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1344
            TabIndex        =   47
            Top             =   1212
            Width           =   2604
         End
         Begin VB.ComboBox cboSystem 
            Height          =   288
            Left            =   1344
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   828
            Width           =   2604
         End
         Begin VB.ComboBox cboEmployee 
            Height          =   288
            Left            =   1344
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   444
            Width           =   2604
         End
         Begin VB.Label lblSystemAssignment 
            Caption         =   "Installed Software"
            Height          =   204
            Index           =   3
            Left            =   4320
            TabIndex        =   48
            Top             =   192
            Width           =   1452
         End
         Begin VB.Label lblSystemAssignment 
            Caption         =   "Serial Number"
            Height          =   204
            Index           =   2
            Left            =   192
            TabIndex        =   44
            Top             =   1248
            Width           =   1068
         End
         Begin VB.Label lblSystemAssignment 
            Caption         =   "System"
            Height          =   204
            Index           =   1
            Left            =   192
            TabIndex        =   43
            Top             =   864
            Width           =   972
         End
         Begin VB.Label lblSystemAssignment 
            Caption         =   "Employee"
            Height          =   204
            Index           =   0
            Left            =   192
            TabIndex        =   42
            Top             =   480
            Width           =   972
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Software"
         Height          =   1260
         Left            =   -74808
         TabIndex        =   36
         Top             =   384
         Width           =   7788
         Begin VB.ComboBox cboSoftwareCategory 
            Height          =   288
            Left            =   1632
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   768
            Width           =   2316
         End
         Begin VB.ComboBox cboSoftware 
            Height          =   288
            Left            =   1632
            TabIndex        =   38
            Top             =   342
            Width           =   2316
         End
         Begin VB.Label lblSoftware 
            Caption         =   "Software Category"
            Height          =   204
            Index           =   1
            Left            =   192
            TabIndex        =   39
            Top             =   816
            Width           =   1356
         End
         Begin VB.Label lblSoftware 
            Caption         =   "Software Title"
            Height          =   204
            Index           =   0
            Left            =   192
            TabIndex        =   37
            Top             =   384
            Width           =   1068
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hardware"
         Height          =   2700
         Left            =   -74808
         TabIndex        =   8
         Top             =   384
         Width           =   7788
         Begin MSComCtl2.DTPicker dtpLeaseExpiration 
            Height          =   288
            Left            =   5472
            TabIndex        =   32
            Top             =   2262
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   508
            _Version        =   393216
            CalendarTitleBackColor=   -2147483635
            CalendarTitleForeColor=   -2147483634
            Format          =   22937601
            CurrentDate     =   36702
         End
         Begin VB.TextBox txtSerialNumber 
            Height          =   288
            Left            =   1632
            TabIndex        =   31
            Top             =   2262
            Width           =   2220
         End
         Begin VB.TextBox txtMonitor 
            Height          =   288
            Left            =   5472
            TabIndex        =   30
            Top             =   1878
            Width           =   2124
         End
         Begin VB.TextBox txtVideoCard 
            Height          =   288
            Left            =   1632
            TabIndex        =   29
            Top             =   1878
            Width           =   2220
         End
         Begin VB.TextBox txtSpeakers 
            Height          =   288
            Left            =   5472
            TabIndex        =   28
            Top             =   1494
            Width           =   2124
         End
         Begin VB.TextBox txtSoundCard 
            Height          =   300
            Left            =   1632
            TabIndex        =   27
            Top             =   1488
            Width           =   2220
         End
         Begin VB.ComboBox cboCD 
            Height          =   288
            Left            =   5472
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1110
            Width           =   972
         End
         Begin VB.TextBox txtHardDrive 
            Height          =   288
            Left            =   1632
            TabIndex        =   25
            Top             =   1110
            Width           =   1356
         End
         Begin VB.TextBox txtMemory 
            Height          =   300
            Left            =   5472
            TabIndex        =   24
            Top             =   720
            Width           =   972
         End
         Begin VB.TextBox txtProcessorSpeed 
            Height          =   300
            Left            =   1632
            TabIndex        =   23
            Top             =   720
            Width           =   1836
         End
         Begin VB.TextBox txtModel 
            Height          =   288
            Left            =   5472
            TabIndex        =   12
            Top             =   342
            Width           =   2220
         End
         Begin VB.ComboBox cboManufacturer 
            Height          =   288
            Left            =   1632
            TabIndex        =   10
            Top             =   342
            Width           =   2220
         End
         Begin VB.Label lblHardware 
            Caption         =   "Lease Expiration"
            Height          =   204
            Index           =   11
            Left            =   4128
            TabIndex        =   22
            Top             =   2304
            Width           =   1260
         End
         Begin VB.Label lblHardware 
            Caption         =   "Serial Number"
            Height          =   204
            Index           =   10
            Left            =   192
            TabIndex        =   21
            Top             =   2304
            Width           =   1164
         End
         Begin VB.Label lblHardware 
            Caption         =   "Monitor"
            Height          =   204
            Index           =   9
            Left            =   4128
            TabIndex        =   20
            Top             =   1920
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Video Card"
            Height          =   204
            Index           =   8
            Left            =   192
            TabIndex        =   19
            Top             =   1920
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Speakers"
            Height          =   204
            Index           =   7
            Left            =   4128
            TabIndex        =   18
            Top             =   1536
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Sound Card"
            Height          =   204
            Index           =   6
            Left            =   192
            TabIndex        =   17
            Top             =   1536
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "CD"
            Height          =   204
            Index           =   5
            Left            =   4128
            TabIndex        =   16
            Top             =   1152
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Hard Drive"
            Height          =   204
            Index           =   4
            Left            =   192
            TabIndex        =   15
            Top             =   1152
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Memory"
            Height          =   204
            Index           =   3
            Left            =   4128
            TabIndex        =   14
            Top             =   768
            Width           =   684
         End
         Begin VB.Label lblHardware 
            Caption         =   "Processor Speed"
            Height          =   204
            Index           =   2
            Left            =   192
            TabIndex        =   13
            Top             =   768
            Width           =   1356
         End
         Begin VB.Label lblHardware 
            Caption         =   "Model"
            Height          =   204
            Index           =   1
            Left            =   4128
            TabIndex        =   11
            Top             =   384
            Width           =   972
         End
         Begin VB.Label lblHardware 
            Caption         =   "Manufacturer"
            Height          =   204
            Index           =   0
            Left            =   192
            TabIndex        =   9
            Top             =   384
            Width           =   972
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Employee"
         Height          =   1740
         Left            =   192
         TabIndex        =   2
         Top             =   384
         Width           =   7788
         Begin VB.ComboBox cboLocation 
            Height          =   288
            Left            =   1440
            TabIndex        =   7
            Top             =   1308
            Width           =   1932
         End
         Begin VB.TextBox txtPhoneNumber 
            Height          =   288
            Left            =   1440
            TabIndex        =   6
            Top             =   828
            Width           =   1452
         End
         Begin VB.TextBox txtFirstName 
            Height          =   288
            Left            =   4512
            TabIndex        =   5
            Top             =   342
            Width           =   1260
         End
         Begin VB.ComboBox cboLastName 
            Height          =   288
            Left            =   1440
            TabIndex        =   4
            Top             =   342
            Width           =   1452
         End
         Begin VB.Label lblEmployee 
            Caption         =   "Location"
            Height          =   204
            Index           =   3
            Left            =   192
            TabIndex        =   35
            Top             =   1344
            Width           =   876
         End
         Begin VB.Label lblEmployee 
            Caption         =   "Phone Number"
            Height          =   204
            Index           =   2
            Left            =   192
            TabIndex        =   34
            Top             =   864
            Width           =   1164
         End
         Begin VB.Label lblEmployee 
            Caption         =   "First Name"
            Height          =   204
            Index           =   1
            Left            =   3552
            TabIndex        =   33
            Top             =   384
            Width           =   876
         End
         Begin VB.Label lblEmployee 
            Caption         =   "Last Name"
            Height          =   204
            Index           =   0
            Left            =   192
            TabIndex        =   3
            Top             =   384
            Width           =   876
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5076
      Width           =   8328
      _ExtentX        =   14690
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12668
            MinWidth        =   1323
            Key             =   "pnlStatus"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9/14/00"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   96
      Top             =   4416
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":037A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":048C
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06B0
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07C2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08D4
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   8328
      _ExtentX        =   14690
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnNew"
            Object.ToolTipText     =   "New Employee"
            ImageKey        =   "New"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "btnNewEmployee"
                  Text            =   "Employee"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "btnNewHardware"
                  Text            =   "Hardware"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "btnNewSoftware"
                  Text            =   "Software"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "btnNewSystemAssignment"
                  Text            =   "System Assignment"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnCut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnCopy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnPaste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnFind"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   8424
      Y1              =   374
      Y2              =   374
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   8424
      Y1              =   364
      Y2              =   364
   End
   Begin VB.Menu zmnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Begin VB.Menu mnuFileNewEmployee 
            Caption         =   "&Employee"
         End
         Begin VB.Menu mnuFileNewHardware 
            Caption         =   "&Hardware"
         End
         Begin VB.Menu mnuFileNewSoftware 
            Caption         =   "&Software"
         End
         Begin VB.Menu mnuFileNewSystemAssignment 
            Caption         =   "System &Assignment"
         End
      End
      Begin VB.Menu zmnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu zmnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu zmnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu zmnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewEmployee 
         Caption         =   "&Employee"
      End
      Begin VB.Menu mnuViewHardware 
         Caption         =   "&Hardware"
      End
      Begin VB.Menu mnuViewSoftware 
         Caption         =   "&Software"
      End
      Begin VB.Menu mnuViewSystemAssignment 
         Caption         =   "System &Assignment"
      End
   End
   Begin VB.Menu zmnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOnlineReports 
         Caption         =   "Online &Reports"
      End
   End
   Begin VB.Menu zmnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare objects
Dim objData As clsSelectData
Dim objInsert As clsInsertData
Dim objRSData As ADODB.Recordset

'Declare variables
Dim arrSerialNumber() As String

Sub LoadEmployees()
   'Setup error handling
   On Error GoTo LoadEmployees_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Loading Employees"
   
   'Declare local variables
   Dim lngRC As Long
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the select data class
   Set objData = New clsSelectData
      
   'Open the recordset
   lngRC = objData.SelectEmployeesAndLocations(g_objConn, g_objRS)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "LoadEmployees", _
         "Call to SelectEmployeesAndLocations failed"
   End If
   
   'Clear combo boxes
   cboLocation.Clear
   cboLastName.Clear
   cboEmployee.Clear
   
   'Load location combo box
   Do While Not g_objRS.EOF
      cboLocation.AddItem g_objRS!Location_Name_VC
      cboLocation.ItemData(cboLocation.NewIndex) = g_objRS!Location_ID
      g_objRS.MoveNext
   Loop
   
   'Get the next recordset
   Set g_objRS = g_objRS.NextRecordset
   
   'Load the last name combo box and employee combo box
   Do While Not g_objRS.EOF
      'Last name combo box
      cboLastName.AddItem g_objRS!Last_Name_VC
      cboLastName.ItemData(cboLastName.NewIndex) = g_objRS!Employee_ID
      'Employee combo box
      cboEmployee.AddItem g_objRS!Last_Name_VC & ", " & g_objRS!First_Name_VC
      cboEmployee.ItemData(cboEmployee.NewIndex) = g_objRS!Employee_ID
      g_objRS.MoveNext
   Loop
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objData = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
LoadEmployees_Err:
   Call ADOError
End Sub
Sub LoadHardware()
   'Setup error handling
   On Error GoTo LoadHardware_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Loading Hardware"
   
   'Declare local variables
   Dim lngRC As Long
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the select data class
   Set objData = New clsSelectData
      
   'Open the recordset
   lngRC = objData.SelectHardwareAndCDs(g_objConn, g_objRS)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "LoadHardware", _
         "Call to SelectHardwareAndCDs failed"
   End If
   
   'Clear combo boxes
   cboCD.Clear
   cboManufacturer.Clear
   
   'Load cd combo box
   Do While Not g_objRS.EOF
      cboCD.AddItem g_objRS!CD_Type_CH
      cboCD.ItemData(cboCD.NewIndex) = g_objRS!CD_ID
      g_objRS.MoveNext
   Loop
   
   'Get the next recordset
   Set g_objRS = g_objRS.NextRecordset
   
   'Load manufacturer combo box
   Do While Not g_objRS.EOF
      'Manufacturer combo box
      cboManufacturer.AddItem g_objRS!Manufacturer_VC
      cboManufacturer.ItemData(cboManufacturer.NewIndex) = g_objRS!Hardware_ID
      g_objRS.MoveNext
   Loop
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objData = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
LoadHardware_Err:
   Call ADOError
End Sub

Sub LoadSoftware()
   'Setup error handling
   On Error GoTo LoadSoftware_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Loading Software"
   
   'Declare local variables
   Dim lngRC As Long
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the select data class
   Set objData = New clsSelectData
      
   'Open the recordset
   lngRC = objData.SelectSoftwareAndCategories(g_objConn, g_objRS)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "LoadSoftware", _
         "Call to SelectSoftwareAndCategories failed"
   End If
   
   'Clear combo and list boxes
   cboSoftware.Clear
   cboSoftwareCategory.Clear
   lstInstalledSoftware.Clear
   
   'Load software combo box and software list
   Do While Not g_objRS.EOF
      'Software combo box
      cboSoftware.AddItem g_objRS!Software_Name_VC
      cboSoftware.ItemData(cboSoftware.NewIndex) = g_objRS!Software_ID
      'Software list
      lstInstalledSoftware.AddItem g_objRS!Software_Name_VC
      lstInstalledSoftware.ItemData(lstInstalledSoftware.NewIndex) = _
         g_objRS!Software_ID
   g_objRS.MoveNext
   Loop
   
   'Get the next recordset
   Set g_objRS = g_objRS.NextRecordset
   
   'Load software category combo box
   Do While Not g_objRS.EOF
      cboSoftwareCategory.AddItem g_objRS!Software_Category_VC
      cboSoftwareCategory.ItemData(cboSoftwareCategory.NewIndex) = _
         g_objRS!Software_Category_ID
      g_objRS.MoveNext
   Loop
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objData = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
LoadSoftware_Err:
   Call ADOError
End Sub
Sub LoadSystems()
   'Setup error handling
   On Error GoTo LoadSystems_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Loading Systems"
   
   'Declare local variables
   Dim lngRC As Long
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the select data class
   Set objData = New clsSelectData
      
   'Open the recordset
   lngRC = objData.SelectSystems(g_objConn, g_objRS)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "LoadSystems", _
         "Call to SelectSystems failed"
   End If
   
   'Clear combo boxes
   cboSystem.Clear
   ReDim arrSerialNumber(0)
   
   'Load system combo box
   Do While Not g_objRS.EOF
      cboSystem.AddItem g_objRS!Manufacturer_VC
      cboSystem.ItemData(cboSystem.NewIndex) = g_objRS!Hardware_ID
      'Re-dimension the array and add new serial number
      ReDim Preserve arrSerialNumber(cboSystem.NewIndex)
      arrSerialNumber(cboSystem.NewIndex) = g_objRS!Serial_Number_VC
   g_objRS.MoveNext
   Loop
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objData = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
LoadSystems_Err:
   Call ADOError
End Sub

Sub ADOError()
   'Declare local variables
   Dim strError
   
   'Loop through the errors collection and display all errors
   For Each g_objError In g_objConn.Errors
      strError = strError & g_objError.Number & " : " & _
         g_objError.Description & vbCrLf & vbCrLf
   Next
   
   'Display the error
   MsgBox strError, vbCritical + vbOKOnly, "ADO Error"
   
   'Unload the form to clean up and then end
   Unload Me
   End
End Sub

Private Sub cboEmployee_Click()
   'Check the ListIndex property
   If cboEmployee.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Setup error handling
   On Error GoTo cboEmployee_Err
   
   'Declare local variables
   Dim lngRC As Long, intIndex As Integer
   
   'Set a reference to the recordset objects
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the select data class
   Set objData = New clsSelectData
      
   'Open the recordset
   lngRC = objData.SelectAssignedSystem(g_objConn, g_objRS, _
      cboEmployee.ItemData(cboEmployee.ListIndex))
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "cboEmployee_Click", _
         "Call to SelectAssignedSystem failed"
   End If
   
   'Select the assigned system and display serial number
   If Not g_objRS.EOF Then
      For intIndex = 0 To cboSystem.ListCount - 1
         If cboSystem.ItemData(intIndex) = g_objRS!Hardware_ID Then
            'Set the index for the assigned system
            cboSystem.ListIndex = intIndex
            'Display the serial number from the array
            txtSystemSerialNumber.Text = arrSerialNumber(intIndex)
            Exit For
         End If
      Next
   End If
   
   'Uncheck all software
   For intIndex = 0 To lstInstalledSoftware.ListCount - 1
      lstInstalledSoftware.Selected(intIndex) = False
   Next
   
   'Check all installed software
   Do While Not g_objRS.EOF
      For intIndex = 0 To lstInstalledSoftware.ListCount - 1
         If lstInstalledSoftware.ItemData(intIndex) = g_objRS!Software_ID Then
            'Check the checkbox for the software title
            lstInstalledSoftware.Selected(intIndex) = True
            Exit For
         End If
      Next
      'Get next software title
      g_objRS.MoveNext
   Loop
   
   'Remove highlights from selected items
   lstInstalledSoftware.ListIndex = -1
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objData = Nothing
   
   'Exit sub
   Exit Sub
   
cboEmployee_Err:
   Call ADOError
End Sub


Private Sub cboLastName_Click()
   'Check the ListIndex property
   If cboLastName.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Declare local variables
   Dim intIndex As Integer
   
   'Move to the first record
   objRSData.MoveFirst
   
   'Find the correct record in the recordset
   objRSData.Find "Employee_ID = " & _
      cboLastName.ItemData(cboLastName.ListIndex)
   
   'Fill the form fields
   txtFirstName.Text = objRSData!First_Name_VC
   txtPhoneNumber.Text = objRSData!Phone_Number_VC
   
   'Loop through the location combo box and find the right entry
   For intIndex = 0 To cboLocation.ListCount - 1
      If cboLocation.ItemData(intIndex) = objRSData!Location_ID Then
         cboLocation.ListIndex = intIndex
         Exit For
      End If
   Next
End Sub


Private Sub cboManufacturer_Click()
   'Check the ListIndex property
   If cboManufacturer.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Declare local variables
   Dim intIndex As Integer
   
   'Move to the first record
   objRSData.MoveFirst
   
   'Find the correct record in the recordset
   objRSData.Find "Hardware_ID = " & _
      cboManufacturer.ItemData(cboManufacturer.ListIndex)
   
   'Fill the form fields
   txtModel.Text = objRSData!Model_VC
   txtProcessorSpeed.Text = objRSData!Processor_Speed_VC
   txtMemory.Text = objRSData!Memory_VC
   txtHardDrive.Text = objRSData!HardDrive_VC
   txtSoundCard.Text = objRSData!Sound_Card_VC
   txtSpeakers.Text = objRSData!Speakers_VC
   txtVideoCard.Text = objRSData!Video_Card_VC
   txtMonitor.Text = objRSData!Monitor_VC
   txtSerialNumber.Text = objRSData!Serial_Number_VC
   dtpLeaseExpiration.Value = objRSData!Lease_Expiration_DT
   
   'Loop through the cd combo box and find the right entry
   For intIndex = 0 To cboCD.ListCount - 1
      If cboCD.ItemData(intIndex) = objRSData!CD_ID Then
         cboCD.ListIndex = intIndex
         Exit For
      End If
   Next
End Sub


Private Sub cboSoftware_Click()
   'Check the ListIndex property
   If cboSoftware.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Declare local variables
   Dim intIndex As Integer
   
   'Move to the first record
   objRSData.MoveFirst
   
   'Find the correct record in the recordset
   objRSData.Find "Software_ID = " & _
      cboSoftware.ItemData(cboSoftware.ListIndex)
   
   'Loop through the software category combo box and find the right entry
   For intIndex = 0 To cboSoftwareCategory.ListCount - 1
      If cboSoftwareCategory.ItemData(intIndex) = _
         objRSData!Software_Category_ID Then
         cboSoftwareCategory.ListIndex = intIndex
         Exit For
      End If
   Next
End Sub


Private Sub cboSystem_Click()
   'Check the ListIndex property
   If cboSystem.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Display the serial number from the array
   txtSystemSerialNumber.Text = arrSerialNumber(cboSystem.ListIndex)
End Sub


Private Sub cmdDelete_Click()

End Sub

Private Sub cmdInsert_Click()
   'Setup error handling
   On Error GoTo cmdInsert_Click_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Processing Insert"
   
   'Declare local variables
   Dim lngRC As Long, strMessage As String, _
      lngInstalledSoftware() As Long, intIndex As Integer, _
      intArrayIndex As Integer

   'Set a reference to the insert data class
   Set objInsert = New clsInsertData
   
   'Process tab data
   Select Case tabData.Tab
   
      Case 0  'Employee
      
         'Insert the new employee
         lngRC = objInsert.InsertEmployee(g_objConn, _
            txtFirstName.Text, _
            cboLastName.Text, _
            txtPhoneNumber.Text, _
            cboLocation.Text, _
            strMessage)
         
         'Ensure we were successful
         If lngRC <> 0 Then
            Err.Raise 513 + vbObjectError, "cmdInsert_Click", _
               "Call to InsertEmployee failed"
         End If
   
         'ReLoad Employees
         Call LoadEmployees
      
         'ReInitialize the tab control so the opened RS is reloaded
         'to reflect the new employee
         Call tabData_Click(0)
      
      Case 1  'Hardware
      
         'Insert the new hardware
         lngRC = objInsert.InsertHardware(g_objConn, _
            cboManufacturer.Text, _
            txtModel.Text, _
            txtProcessorSpeed.Text, _
            txtMemory.Text, _
            txtHardDrive.Text, _
            txtSoundCard.Text, _
            txtSpeakers.Text, _
            txtVideoCard.Text, _
            txtMonitor.Text, _
            txtSerialNumber.Text, _
            dtpLeaseExpiration.Value, _
            cboCD.ItemData(cboCD.ListIndex), _
            strMessage)
         
         'Ensure we were successful
         If lngRC <> 0 Then
            Err.Raise 513 + vbObjectError, "cmdInsert_Click", _
               "Call to InsertHardware failed"
         End If
   
         'ReLoad Hardware
         Call LoadHardware
         
         'Reload Systems combo box on the System Assignment tab
         Call LoadSystems
      
         'ReInitialize the tab control so the opened RS is reloaded
         'to reflect the new hardware
         Call tabData_Click(1)
      
      Case 2  'Software
      
         'Insert the new software
         lngRC = objInsert.InsertSoftware(g_objConn, _
            cboSoftware.Text, _
            cboSoftwareCategory.ItemData(cboSoftwareCategory.ListIndex), _
            strMessage)
         
         'Ensure we were successful
         If lngRC <> 0 Then
            Err.Raise 513 + vbObjectError, "cmdInsert_Click", _
               "Call to InsertSoftware failed"
         End If
   
         'ReLoad Software
         Call LoadSoftware
      
         'ReInitialize the tab control so the opened RS is reloaded
         'to reflect the new software
         Call tabData_Click(2)
      
      Case 3  'System Assignment
      
         'Build an array of installed software
         intArrayIndex = -1
         For intIndex = 0 To lstInstalledSoftware.ListCount - 1
            'If the selected item is checked then
            If lstInstalledSoftware.Selected(intIndex) = True Then
               'Increment the array index variable and redim the array
               intArrayIndex = intArrayIndex + 1
               ReDim Preserve lngInstalledSoftware(intArrayIndex)
               'Add the ItemData property value to the array
               lngInstalledSoftware(intArrayIndex) = _
                  lstInstalledSoftware.ItemData(intIndex)
            End If
         Next
         
         'Insert the system assignment and software
         lngRC = objInsert.InsertSystemAssignment(g_objConn, _
            cboEmployee.ItemData(cboEmployee.ListIndex), _
            cboSystem.ItemData(cboSystem.ListIndex), _
            lngInstalledSoftware(), _
            strMessage)
         
         'Ensure we were successful
         If lngRC <> 0 Then
            Err.Raise 513 + vbObjectError, "cmdInsert_Click", _
               "Call to InsertSystemAssignment failed"
         End If
   
         'ReLoad Systems
         Call LoadSystems
      
         'ReInitialize the tab control so the opened RS is reloaded
         'to reflect the new system assignment
         Call tabData_Click(3)
      
   End Select
   
   'Remove reference to object
   Set objInsert = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
cmdInsert_Click_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub

Private Sub cmdUpdate_Click()

End Sub


Private Sub Form_Load()
   'Display the login form
   frmLogin.Show vbModal
   
   'Show the form and set the mouse pointer to busy
   Me.Show
   Me.MousePointer = vbArrowHourglass
   DoEvents
   
   'Load Employees
   Call LoadEmployees
   
   'Load Hardware
   Call LoadHardware
   
   'Load Software
   Call LoadSoftware
   
   'Load Systems
   Call LoadSystems
   
   'Initialize controls
   tabData.Tab = 0
   Call tabData_Click(0)
   
   'Return mouse pointer to default value
   Me.MousePointer = vbDefault
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Dereference local objects
   If Not objRSData Is Nothing Then
      If objRSData.State = adStateOpen Then
         objRSData.Close
      End If
      Set objRSData = Nothing
   End If
   
   Set objData = Nothing
   Set objInsert = Nothing
   
   'Termination the database connection
   Call TerminateConnection
End Sub


Private Sub mnuFileExit_Click()
   Unload Me
End Sub


Private Sub mnuFileNewEmployee_Click()
   'Clear the fields on the Employee tab
   txtFirstName.Text = Empty
   cboLastName.ListIndex = -1
   txtPhoneNumber.Text = Empty
   cboLocation.ListIndex = -1
End Sub

Private Sub mnuFileNewHardware_Click()
   'Clear the fields on the Hardware tab
   cboManufacturer.ListIndex = -1
   txtModel.Text = Empty
   txtProcessorSpeed.Text = Empty
   txtMemory.Text = Empty
   txtHardDrive.Text = Empty
   txtSoundCard.Text = Empty
   txtSpeakers.Text = Empty
   txtVideoCard.Text = Empty
   txtMonitor.Text = Empty
   txtSerialNumber.Text = Empty
   dtpLeaseExpiration.Value = Now
   cboCD.ListIndex = -1
End Sub

Private Sub mnuFileNewSoftware_Click()
   'Clear the fields on the Software tab
   cboSoftware.ListIndex = -1
   cboSoftwareCategory.ListIndex = -1
End Sub


Private Sub mnuFileNewSystemAssignment_Click()
   'Declare variables
   Dim intIndex As Integer
   
   'Clear the fields on the System Assignment tab
   cboEmployee.ListIndex = -1
   cboSystem.ListIndex = -1
   txtSystemSerialNumber.Text = Empty
   
   'Uncheck all software
   For intIndex = 0 To lstInstalledSoftware.ListCount - 1
      lstInstalledSoftware.Selected(intIndex) = False
   Next
End Sub


Private Sub mnuViewEmployee_Click()
   'Make the employee tab active
   tabData.Tab = 0
End Sub

Private Sub mnuViewHardware_Click()
   'Make the hardware tab active
   tabData.Tab = 1
End Sub


Private Sub mnuViewSoftware_Click()
   'Make the software tab active
   tabData.Tab = 2
End Sub

Private Sub mnuViewSystemAssignment_Click()
   'Make the system assignment tab active
   tabData.Tab = 3
End Sub


Private Sub tabData_Click(PreviousTab As Integer)
   'Setup error handling
   On Error GoTo tabData_Click_Err
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Processing Tab Data"
   
   'Declare local variables
   Dim lngRC As Long, strSQL As String
   
   'Close and dereference recordset if its already open
   If Not objRSData Is Nothing Then
      If objRSData.State = adStateOpen Then
         objRSData.Close
      End If
      Set objRSData = Nothing
   End If
   
   'Set a reference to the recordset object
   Set objRSData = New ADODB.Recordset
   
   'Set the appropriate string
   Select Case tabData.Tab
      Case 0  'Employee
         strSQL = "up_select_employees"
      Case 1  'Hardware
         strSQL = "up_select_hardware"
      Case 2  'Software
         strSQL = "up_select_software"
      Case 3  'System Assignment
   End Select
   
   'If one of the first three tabs, open a recordset
   If tabData.Tab <> 3 Then
   
      'Set a reference to the select data class
      Set objData = New clsSelectData
         
      'Open the recordset
      lngRC = objData.ExecuteSQL(g_objConn, objRSData, strSQL, adCmdStoredProc)
      
      'Ensure we were successful
      If lngRC <> 0 Then
         Err.Raise 513 + vbObjectError, "tabData_Click", _
            "Call to ExecuteSQL failed"
      End If
      
   End If

   'Show the first record if available
   Select Case tabData.Tab
      Case 0  'Employee
         If cboLastName.ListCount <> -1 Then
            cboLastName.ListIndex = 0
         End If
      Case 1  'Hardware
         If cboManufacturer.ListCount <> -1 Then
            cboManufacturer.ListIndex = 0
         End If
      Case 2  'Software
         If cboSoftware.ListCount <> -1 Then
            cboSoftware.ListIndex = 0
         End If
      Case 3  'System Assignment
         If cboEmployee.ListCount <> -1 Then
            'The Style property is set to 2 - Dropdown List and
            'will not fire the click event if the ListIndex property
            'is already set to the same number we are using here.
            'To circumvent this problem, we first set the ListIndex
            'property to -1 and then to 0 to ensure the click event
            'gets fired
            cboEmployee.ListIndex = -1
            cboEmployee.ListIndex = 0
         End If
   End Select
   
   'Remove references to objects
   Set objData = Nothing
   
   'Display status
   StatusBar1.Panels("pnlStatus").Text = "Ready"
   
   'Exit sub
   Exit Sub
   
tabData_Click_Err:
   Call ADOError
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   'Process the appropriate toolbar button and call the
   'corresponding menu item
   Select Case Button.Key
   
      Case "btnNew"
         Call mnuFileNewEmployee_Click
         
      Case "btnCut"
      
      Case "btnCopy"
      
      Case "btnPaste"
      
      Case "btnFind"
      
      Case "btnHelp"
      
   End Select
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   'Process the appropriate button menu item and call the
   'corresponding menu item
   Select Case ButtonMenu.Key
   
      Case "btnNewEmployee"
         Call mnuFileNewEmployee_Click
         
      Case "btnNewHardware"
         Call mnuFileNewHardware_Click
         
      Case "btnNewSoftware"
         Call mnuFileNewSoftware_Click
         
      Case "btnNewSystemAssignment"
         Call mnuFileNewSystemAssignment_Click
         
   End Select
End Sub


