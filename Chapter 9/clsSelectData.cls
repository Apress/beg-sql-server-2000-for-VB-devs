VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SelectEmployeesAndLocations( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset) As Long
   
   'Setup error handling
   On Error GoTo SelectEmployeesAndLocations_Err
   
   'Open the recordset object
   objRS.Open "up_select_employees_and_locations", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectEmployeesAndLocations = 0
   
   'Exit Function
   Exit Function
   
SelectEmployeesAndLocations_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "SelectEmployeesAndLocations")
   SelectEmployeesAndLocations = 1
End Function

Public Function SelectSystems( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset) As Long

   'Setup error handling
   On Error GoTo SelectSystems_Err
   
   'Open the recordset object
   objRS.Open "up_select_unassigned_and_assigned_hardware", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectSystems = 0
   
   'Exit function
   Exit Function
   
SelectSystems_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "SelectSystems")
   SelectSystems = 1
End Function

Public Function SelectAssignedSystem( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset, _
   ByVal lngEmployee As Long) As Long

   'Setup error handling
   On Error GoTo SelectAssignedSystem_Err
   
   'Open the recordset object
   objRS.Open "up_parmsel_assigned_system (" & lngEmployee & ")", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectAssignedSystem = 0
   
   'Exit function
   Exit Function
   
SelectAssignedSystem_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "SelectAssignedSystem")
   SelectAssignedSystem = 1
End Function

Public Function ExecuteSQL( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset, _
   ByVal strSQL As String, _
   ByVal lngCommandType As CommandTypeEnum) As Long

   'Setup error handling
   On Error GoTo ExecuteSQL_Err
   
   'Use a client side cursor
   objRS.CursorLocation = adUseClient
   
   'Open the recordset object
   objRS.Open strSQL, _
      objConn, adOpenStatic, adLockReadOnly, lngCommandType
   
   'Disconnect the recordset
   Set objRS.ActiveConnection = Nothing
   
   'Set the return code
   ExecuteSQL = 0
   
   'Exit function
   Exit Function
   
ExecuteSQL_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "ExecuteSQL")
   ExecuteSQL = 1
End Function

Public Function SelectSoftwareAndCategories( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset) As Long

   'Setup error handling
   On Error GoTo SelectSoftwareAndCategories_Err
   
   'Open the recordset object
   objRS.Open "up_select_software_and_categories", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectSoftwareAndCategories = 0
   
   'Exit function
   Exit Function
   
SelectSoftwareAndCategories_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "SelectSoftwareAndCategories")
   SelectSoftwareAndCategories = 1
End Function

Public Function SelectHardwareAndCDs( _
   ByRef objConn As ADODB.Connection, _
   ByRef objRS As ADODB.Recordset) As Long

   'Setup error handling
   On Error GoTo SelectHardwareAndCDs_Err
   
   'Open the recordset object
   objRS.Open "up_select_hardware_and_cds", _
      objConn, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
   
   'Set the return code
   SelectHardwareAndCDs = 0
   
   'Exit function
   Exit Function
   
SelectHardwareAndCDs_Err:
   'Call the error handler and return a return code of 1
   Call ADOError(objConn, "SelectHardwareAndCDs")
   SelectHardwareAndCDs = 1
End Function



Private Sub ADOError(ByRef objConn As ADODB.Connection, _
   ByVal strProcedure As String)
   
   'Declare local objects
   Dim objError As ADODB.Error
   
   'Loop through the errors collection and log all errors
   For Each objError In objConn.Errors
      App.LogEvent "The following error was encountered in the " & _
         strProcedure & " procedure:" & vbCrLf & _
         objError.Number & " : " & objError.Description, _
         vbLogEventTypeError
   Next
   
   'Dereference error object
   Set objError = Nothing
End Sub
