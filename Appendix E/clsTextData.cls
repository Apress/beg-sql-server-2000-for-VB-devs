VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelectNotes( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset, _
   ByVal lngHardwareID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo SelectNotes_Err
   
   'Open the recordset object
   objRS.Open "up_parmsel_system_notes (" & lngHardwareID & ")", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectNotes = 0
   
   'Exit function
   Exit Function
   
SelectNotes_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   SelectNotes = 1
End Function
Public Function SelectPartialNotes( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset, _
   ByVal lngNotesID As Long, _
   ByVal lngOffset As Long, _
   ByVal lngLength As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo SelectPartialNotes_Err
   
   'Open the recordset object
   objRS.Open "up_parmsel_read_partial_system_notes (" & _
      lngNotesID & "," & lngOffset & "," & lngLength & ")", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectPartialNotes = 0
   
   'Exit function
   Exit Function
   
SelectPartialNotes_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   SelectPartialNotes = 1
End Function

Public Function InsertNotes( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngHardwareID As Long, _
   ByRef strNotes As String, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo InsertNotes_Err
   
   'Declare local variables and objects
   Dim lngTextSize As Long, lngOffset As Long, lngChunkSize As Long
   Dim objCmd As ADODB.Command
   Dim objParm As ADODB.Parameter
   
   'Validate notes data
   If RTrim(Len(strNotes)) = 0 Then
      strMessage = "Notes data is zero length" & vbCrLf & vbCrLf
      Err.Raise 513 + vbObjectError, "InsertNotes", _
         "Data validation failed"
   End If
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmins_system_notes"
   objCmd.CommandType = adCmdStoredProc
   
   '************************************************
   'First parameter - Return value
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Name = "RC"
   objParm.Type = adInteger
   objParm.Direction = adParamReturnValue
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
      
   '************************************************
   'Second parameter - Hardware ID
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Name = "HardwareID"
   objParm.Type = adInteger
   objParm.Direction = adParamInput
   objParm.Value = lngHardwareID
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
   
   '************************************************
   'Third parameter - Notes data
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Attributes = adFldLong
   objParm.Name = "Notes"
   objParm.Type = adLongVarChar
   objParm.Direction = adParamInput
   objParm.Size = Len(strNotes)
   
   'Set default values
   lngTextSize = Len(strNotes)
   lngChunkSize = (lngTextSize / 10)
   lngOffset = 1
   
   'Loop through the string and append the data to the parameter object
   Do While lngOffset < lngTextSize
      objParm.AppendChunk Mid(strNotes, lngOffset, lngChunkSize)
      lngOffset = lngOffset + lngChunkSize
   Loop
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "InsertNotes", _
         "up_parmins_system_notes failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   InsertNotes = 0
   
   'Exit function
   Exit Function
   
InsertNotes_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   InsertNotes = 1
End Function
Public Function ReplaceNotes( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngNotesID As Long, _
   ByRef strNotes As String, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo ReplaceNotes_Err
   
   'Declare local variables and objects
   Dim lngTextSize As Long, lngOffset As Long, lngChunkSize As Long
   Dim objCmd As ADODB.Command
   Dim objParm As ADODB.Parameter
   
   'Validate notes data
   If RTrim(Len(strNotes)) = 0 Then
      strMessage = "Notes data is zero length" & vbCrLf & vbCrLf
      Err.Raise 513 + vbObjectError, "ReplaceNotes", _
         "Data validation failed"
   End If
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_replace_system_notes"
   objCmd.CommandType = adCmdStoredProc
   
   '************************************************
   'First parameter - Return value
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Name = "RC"
   objParm.Type = adInteger
   objParm.Direction = adParamReturnValue
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
      
   '************************************************
   'Second parameter - Notes ID
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Name = "NotesID"
   objParm.Type = adInteger
   objParm.Direction = adParamInput
   objParm.Value = lngNotesID
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
   
   '************************************************
   'Third parameter - Notes data
   '************************************************
   'Set a reference to the parameter object
   Set objParm = New ADODB.Parameter
   
   'Set the parameter properties
   objParm.Attributes = adFldLong
   objParm.Name = "Notes"
   objParm.Type = adLongVarChar
   objParm.Direction = adParamInput
   objParm.Size = Len(strNotes)
   
   'Set default values
   lngTextSize = Len(strNotes)
   lngChunkSize = (lngTextSize / 10)
   lngOffset = 1
   
   'Loop through the string and append the data to the parameter object
   Do While lngOffset < lngTextSize
      objParm.AppendChunk Mid(strNotes, lngOffset, lngChunkSize)
      lngOffset = lngOffset + lngChunkSize
   Loop
   
   'Append the parameter to the parameters collection
   objCmd.Parameters.Append objParm
   
   'Remove reference to the parameter object
   Set objParm = Nothing
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "ReplaceNotes", _
         "up_parmupd_replace_system_notes failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   ReplaceNotes = 0
   
   'Exit function
   Exit Function
   
ReplaceNotes_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   ReplaceNotes = 1
End Function
Public Function UpdateNotes( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngNotesID As Long, _
   ByVal lngOffset As Long, _
   ByVal lngLength As Long, _
   ByRef strNotes As String, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo UpdateNotes_Err
   
   'Declare local objects
   Dim objCmd As ADODB.Command
   
   'Validate notes data
   If RTrim(Len(strNotes)) = 0 Then
      strMessage = "Notes data is zero length" & vbCrLf & vbCrLf
      Err.Raise 513 + vbObjectError, "UpdateNotes", _
         "Data validation failed"
   End If
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_update_system_notes"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("NotesID", _
      adInteger, adParamInput, , lngNotesID)
   objCmd.Parameters.Append objCmd.CreateParameter("Offset", _
      adInteger, adParamInput, , lngOffset)
   objCmd.Parameters.Append objCmd.CreateParameter("Length", _
      adInteger, adParamInput, , lngLength)
   objCmd.Parameters.Append objCmd.CreateParameter("Notes", _
      adLongVarChar, adParamInput, Len(strNotes), strNotes)
    
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "UpdateNotes", _
         "up_parmupd_update_system_notes failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   UpdateNotes = 0
   
   'Exit function
   Exit Function
   
UpdateNotes_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   UpdateNotes = 1
End Function




