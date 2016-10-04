VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeleteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function DeleteSystemAssignment( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngAssignmentID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo DeleteSystemAssignment_Err
   
   'Declare local objects
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmdel_system_assignment"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("SystemID", _
      adInteger, adParamInput, , lngAssignmentID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "DeleteSystemAssignment", _
         "up_parmdel_system_assignment failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   DeleteSystemAssignment = 0
   
   'Exit function
   Exit Function
   
DeleteSystemAssignment_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   DeleteSystemAssignment = 1
End Function



Public Function DeleteSoftware( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngSoftwareID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo DeleteSoftware_Err
   
   'Declare local objects
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmdel_software"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("SoftwareID", _
      adInteger, adParamInput, , lngSoftwareID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "DeleteSoftware", _
         "up_parmdel_software failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
  
   'Set the return code
   DeleteSoftware = 0
   
   'Exit function
   Exit Function
   
DeleteSoftware_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   DeleteSoftware = 1
End Function

Public Function DeleteHardware( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngHardwareID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo DeleteHardware_Err
   
   'Declare local objects
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmdel_hardware"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("HardwareID", _
      adInteger, adParamInput, , lngHardwareID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "DeleteHardware", _
         "up_parmdel_hardware failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
  
   'Set the return code
   DeleteHardware = 0
   
   'Exit function
   Exit Function
   
DeleteHardware_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   DeleteHardware = 1
End Function

Public Function DeleteEmployee( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngEmployeeID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo DeleteEmployee_Err
   
   'Declare local objects
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmdel_employee"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("EmployeeID", _
      adInteger, adParamInput, , lngEmployeeID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "DeleteEmployee", _
         "up_parmdel_employee failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   DeleteEmployee = 0
   
   'Exit function
   Exit Function
   
DeleteEmployee_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   DeleteEmployee = 1
End Function


