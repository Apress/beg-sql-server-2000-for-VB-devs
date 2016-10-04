VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notes"
   ClientHeight    =   4032
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6912
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032
   ScaleWidth      =   6912
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   5856
      TabIndex        =   10
      Text            =   "0"
      Top             =   2400
      Width           =   492
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   4992
      TabIndex        =   9
      Text            =   "0"
      Top             =   2400
      Width           =   492
   End
   Begin VB.CommandButton cmdReplaceNotes 
      Caption         =   "Replace Notes"
      Height          =   396
      Left            =   4992
      TabIndex        =   6
      Top             =   1536
      Width           =   1836
   End
   Begin VB.CommandButton cmdUpdatePartialNotes 
      Caption         =   "Update Partial Notes"
      Height          =   396
      Left            =   4992
      TabIndex        =   5
      Top             =   1056
      Width           =   1836
   End
   Begin VB.CommandButton cmdInsertNewNotes 
      Caption         =   "Insert New Notes"
      Height          =   396
      Left            =   4992
      TabIndex        =   4
      Top             =   576
      Width           =   1836
   End
   Begin VB.CommandButton cmdReadPartialNotes 
      Caption         =   "Read Partial Notes"
      Height          =   396
      Left            =   4992
      TabIndex        =   3
      Top             =   96
      Width           =   1836
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3852
      Left            =   96
      ScaleHeight     =   3804
      ScaleWidth      =   4764
      TabIndex        =   0
      Top             =   96
      Width           =   4812
      Begin VB.TextBox txtNotes 
         BorderStyle     =   0  'None
         Height          =   3180
         Left            =   10
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   576
         Width           =   4716
      End
      Begin VB.Label lblSystemNotes 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "System Notes for "
         Height          =   204
         Left            =   672
         TabIndex        =   1
         Top             =   192
         Width           =   3948
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   96
         Picture         =   "frmNotes.frx":000C
         Top             =   96
         Width           =   384
      End
   End
   Begin VB.Label lblMessage 
      Height          =   204
      Left            =   4992
      TabIndex        =   11
      Top             =   3744
      Width           =   1836
   End
   Begin VB.Label Label2 
      Caption         =   "Length"
      Height          =   204
      Left            =   5856
      TabIndex        =   8
      Top             =   2112
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Offset"
      Height          =   204
      Left            =   4992
      TabIndex        =   7
      Top             =   2112
      Width           =   492
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare variables
Dim lngHardwareID As Long
Dim lngNotesID As Long

'Declare objects
Dim objText As clsTextData


Public Sub DisplayNotes( _
   ByVal lngSystemID As Long, _
   ByVal strSystem As String)
   
   'Setup error handling
   On Error GoTo DisplayNotes_Err
   
   'Declare local variables
   Dim lngRC As Long, lngTextSize As Long, lngOffset As Long, _
      lngChunkSize As Long, strMessage As String
   
   'Trim string if necessary
   If Left$(strSystem, 4) = "NA -" Then
      strSystem = Right$(strSystem, Len(strSystem) - 5)
   End If
   
   'Display system manufacturer and model
   lblSystemNotes.Caption = lblSystemNotes.Caption & strSystem
   
   'Save system id
   lngHardwareID = lngSystemID
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the text data class
   Set objText = New clsTextData
      
   'Open the recordset
   lngRC = objText.SelectNotes( _
      g_objConn, _
      g_objRS, _
      lngHardwareID, _
      strMessage)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "DisplayNotes", _
         "Call to SelectNotes failed"
   End If
   
   'Load notes text box and save the notes id
   If Not g_objRS.EOF Then
   
      'Get the actual size of the text field in the RS
      lngTextSize = g_objRS!Hardware_Notes_TX.ActualSize
      
      'Set the chunk size to be retrieved
      lngChunkSize = 102400
      
      If lngTextSize > lngChunkSize Then
      
         'Process data in chunks
         Do While lngOffset <= lngTextSize
         
            'Get a chunk of data and add it to the notes text box
            txtNotes.Text = txtNotes.Text & _
               g_objRS!Hardware_Notes_TX.GetChunk(lngChunkSize)
               
            'Increment the offset counter
            lngOffset = lngOffset + lngChunkSize
            
         Loop
         
      Else
      
         'Small amount of data, just load the text box
         txtNotes.Text = g_objRS!Hardware_Notes_TX
         
      End If
      
      'Save the notes id
      lngNotesID = g_objRS!Hardware_Notes_ID
      
   End If
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objText = Nothing
   
   'Display the form as modal
   Me.Show vbModal
   
   'Exit sub
   Exit Sub
   
DisplayNotes_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub


Private Sub cmdInsertNewNotes_Click()
   'Setup error handling
   On Error GoTo cmdInsertNewNotes_Click_Err
   
   'Declare local variables
   Dim lngRC As Long, strMessage As String
   
   'Set a reference to the text data class
   Set objText = New clsTextData
      
   'Insert the notes
   lngRC = objText.InsertNotes( _
      g_objConn, _
      lngHardwareID, _
      txtNotes.Text, _
      strMessage)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "cmdInsertNewNotes_Click", _
         "Call to InsertNotes failed"
   End If

   'Remove references to objects
   Set objText = Nothing
   
   'Display message
   lblMessage.Caption = "Insert Successful"
   
   'Exit sub
   Exit Sub
   
cmdInsertNewNotes_Click_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub

Private Sub cmdReadPartialNotes_Click()
   'Setup error handling
   On Error GoTo cmdReadPartialNotes_Click_Err
   
   'Declare local variables
   Dim lngRC As Long, strMessage As String
   
   'Set a reference to the recordset object
   Set g_objRS = New ADODB.Recordset
   
   'Set a reference to the text data class
   Set objText = New clsTextData
      
   'Open the recordset
   lngRC = objText.SelectPartialNotes( _
      g_objConn, _
      g_objRS, _
      lngNotesID, _
      txtOffset.Text, _
      txtLength.Text, _
      strMessage)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "cmdReadPartialNotes_Click", _
         "Call to SelectPartialNotes failed"
   End If

   'Load the notes text box
   txtNotes.Text = g_objRS!Hardware_Notes_TX
   
   'Close the recordset
   g_objRS.Close
   
   'Remove references to objects
   Set g_objRS = Nothing
   Set objText = Nothing
   
   'Clear any previous messages
   lblMessage.Caption = Empty
   
   'Exit sub
   Exit Sub
   
cmdReadPartialNotes_Click_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub


Private Sub cmdReplaceNotes_Click()
   'Setup error handling
   On Error GoTo cmdReplaceNotes_Click_Err
   
   'Declare local variables
   Dim lngRC As Long, strMessage As String
   
   'Set a reference to the text data class
   Set objText = New clsTextData
      
   'Replace existing notes
   lngRC = objText.ReplaceNotes( _
      g_objConn, _
      lngNotesID, _
      txtNotes.Text, _
      strMessage)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "cmdReplaceNotes_Click", _
         "Call to ReplaceNotes failed"
   End If

   'Remove references to objects
   Set objText = Nothing
   
   'Display message
   lblMessage.Caption = "Replacement Successful"
   
   'Exit sub
   Exit Sub
   
cmdReplaceNotes_Click_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub

Private Sub cmdUpdatePartialNotes_Click()
   'Setup error handling
   On Error GoTo cmdUpdatePartialNotes_Click_Err
   
   'Declare local variables
   Dim lngRC As Long, strMessage As String
   
   'Set a reference to the text data class
   Set objText = New clsTextData
      
   'Update a portion of the notes
   lngRC = objText.UpdateNotes( _
      g_objConn, _
      lngNotesID, _
      txtOffset.Text, _
      txtLength.Text, _
      txtNotes.SelText, _
      strMessage)
   
   'Ensure we were successful
   If lngRC <> 0 Then
      Err.Raise 513 + vbObjectError, "cmdUpdatePartialNotes_Click", _
         "Call to UpdateNotes failed"
   End If

   'Remove references to objects
   Set objText = Nothing
   
   'Display message
   lblMessage.Caption = "Update Successful"
   
   'Exit sub
   Exit Sub
   
cmdUpdatePartialNotes_Click_Err:
   'Display errors from the function call
   MsgBox strMessage, vbCritical + vbOKOnly, "Hardware Tracking"
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Remove reference to local objects
   Set objText = Nothing
End Sub


