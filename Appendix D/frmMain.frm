VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
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
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   420
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Employee"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Hardware"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Software"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "System Assignment"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "System Assignment"
         Height          =   3276
         Left            =   192
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
            Format          =   77004801
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
         Left            =   -74808
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
            Object.Width           =   12584
            MinWidth        =   1323
            Key             =   "pnlStatus"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "09/10/2000"
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

