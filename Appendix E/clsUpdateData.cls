VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function UpdateEmployee( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngEmployeeID As Long, _
   ByVal strFirstName As String, _
   ByVal strLastName As String, _
   ByVal strPhoneNumber As String, _
   ByVal lngLocationID As Long, _
   ByVal strLocation As String, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo UpdateEmployee_Err
   
   'Declare local variables and objects
   Dim blnValidated As Boolean
   Dim objCmd As ADODB.Command
   
   'Set default values
   blnValidated = True
   
   'Validate First Name
   If RTrim(Len(strFirstName)) = 0 Then
      strMessage = "First Name is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strFirstName)) > 15 Then
      strMessage = "First Name is greater than 15 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
  
   'Validate Last Name
   If RTrim(Len(strLastName)) = 0 Then
      strMessage = strMessage & "Last Name is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strLastName)) > 15 Then
      strMessage = strMessage & _
         "Last Name is greater than 15 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
  
   'Validate Phone Number
   If RTrim(Len(strPhoneNumber)) = 0 Then
      strMessage = strMessage & "Phone Number is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strPhoneNumber)) > 20 Then
      strMessage = strMessage & _
         "Phone Number is greater than 20 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
  
   'Validate Location
   If RTrim(Len(strLocation)) = 0 Then
      strMessage = strMessage & "Location is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strLocation)) > 30 Then
      strMessage = strMessage & _
         "Location is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
  
   'Check validation variable
   If Not blnValidated Then
      Err.Raise 513 + vbObjectError, "UpdateEmployee", _
         "Data validation failed"
   End If
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_employee"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("EmployeeID", _
      adInteger, adParamInput, , lngEmployeeID)
   objCmd.Parameters.Append objCmd.CreateParameter("FirstName", _
      adVarChar, adParamInput, 15, strFirstName)
   objCmd.Parameters.Append objCmd.CreateParameter("LastName", _
      adVarChar, adParamInput, 15, strLastName)
   objCmd.Parameters.Append objCmd.CreateParameter("PhoneNumber", _
      adVarChar, adParamInput, 20, strPhoneNumber)
   objCmd.Parameters.Append objCmd.CreateParameter("LocationID", _
      adInteger, adParamInput, , lngLocationID)
   objCmd.Parameters.Append objCmd.CreateParameter("Location", _
      adVarChar, adParamInput, 30, strLocation)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "UpdateEmployee", _
         "up_parmupd_employee failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Set the return code
   UpdateEmployee = 0
   
   'Exit function
   Exit Function
   
UpdateEmployee_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   UpdateEmployee = 1
End Function
Public Function UpdateHardware( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngHardwareID As Long, _
   ByVal strManufacturer As String, _
   ByVal strModel As String, _
   ByVal strProcessorSpeed As String, _
   ByVal strMemory As String, _
   ByVal strHardDrive As String, _
   ByVal strSoundCard As String, _
   ByVal strSpeakers As String, _
   ByVal strVideoCard As String, _
   ByVal strMonitor As String, _
   ByVal strSerialNumber As String, _
   ByVal dteLeaseExpiration As Date, _
   ByVal lngCDID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo UpdateHardware_Err
   
   'Declare local variables and objects
   Dim blnValidated As Boolean
   Dim objCmd As ADODB.Command
   
   'Set default values
   blnValidated = True
   
   'Validate Manufacturer
   If RTrim(Len(strManufacturer)) = 0 Then
      strMessage = "Manufacturer is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strManufacturer)) > 30 Then
      strMessage = "Manufacturer is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Model
   If RTrim(Len(strModel)) = 0 Then
      strMessage = "Model is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strModel)) > 30 Then
      strMessage = "Model is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Processor Speed
   If RTrim(Len(strProcessorSpeed)) = 0 Then
      strMessage = "Processor Speed is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strProcessorSpeed)) > 20 Then
      strMessage = "Processor Speed is greater than 20 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Memory
   If RTrim(Len(strMemory)) = 0 Then
      strMessage = "Memory is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strMemory)) > 10 Then
      strMessage = "Memory is greater than 10 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Hard Drive
   If RTrim(Len(strHardDrive)) = 0 Then
      strMessage = "Hard Drive is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strHardDrive)) > 15 Then
      strMessage = "Hard Drive is greater than 15 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Sound Card
   If RTrim(Len(strSoundCard)) = 0 Then
      strMessage = "Sound Card is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strSoundCard)) > 30 Then
      strMessage = "Sound Card is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Speakers
   If RTrim(Len(strSpeakers)) = 0 Then
      strMessage = "Speakers is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strSpeakers)) > 30 Then
      strMessage = "Speakers is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Video Card
   If RTrim(Len(strVideoCard)) = 0 Then
      strMessage = "Video Card is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strVideoCard)) > 30 Then
      strMessage = "Video Card is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Monitor
   If RTrim(Len(strMonitor)) = 0 Then
      strMessage = "Monitor is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strMonitor)) > 30 Then
      strMessage = "Monitor is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Serial Number
   If RTrim(Len(strSerialNumber)) = 0 Then
      strMessage = "Serial Number is zero length" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   If RTrim(Len(strSerialNumber)) > 30 Then
      strMessage = "Serial Number is greater than 30 characters" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Validate Lease Expiration
   If dteLeaseExpiration <= Now Then
      strMessage = "Lease Expiration is not a date in the future" & vbCrLf & vbCrLf
      blnValidated = False
   End If
   
   'Check validation variable
   If Not blnValidated Then
      Err.Raise 513 + vbObjectError, "UpdateHardware", _
         "Data validation failed"
   End If
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_hardware"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("HardwareID", _
      adInteger, adParamInput, , lngHardwareID)
   objCmd.Parameters.Append objCmd.CreateParameter("Manufacturer", _
      adVarChar, adParamInput, 30, strManufacturer)
   objCmd.Parameters.Append objCmd.CreateParameter("Model", _
      adVarChar, adParamInput, 30, strModel)
   objCmd.Parameters.Append objCmd.CreateParameter("Processor", _
      adVarChar, adParamInput, 20, strProcessorSpeed)
   objCmd.Parameters.Append objCmd.CreateParameter("Memory", _
      adVarChar, adParamInput, 10, strMemory)
   objCmd.Parameters.Append objCmd.CreateParameter("HardDrive", _
      adVarChar, adParamInput, 15, strHardDrive)
   objCmd.Parameters.Append objCmd.CreateParameter("SoundCard", _
      adVarChar, adParamInput, 30, strSoundCard)
   objCmd.Parameters.Append objCmd.CreateParameter("Speakers", _
      adVarChar, adParamInput, 30, strSpeakers)
   objCmd.Parameters.Append objCmd.CreateParameter("VideoCard", _
      adVarChar, adParamInput, 30, strVideoCard)
   objCmd.Parameters.Append objCmd.CreateParameter("Monitor", _
      adVarChar, adParamInput, 30, strMonitor)
   objCmd.Parameters.Append objCmd.CreateParameter("SerialNumber", _
      adVarChar, adParamInput, 30, strSerialNumber)
   objCmd.Parameters.Append objCmd.CreateParameter("LeaseExp", _
      adVarChar, adParamInput, 22, dteLeaseExpiration)
   objCmd.Parameters.Append objCmd.CreateParameter("CDID", _
      adInteger, adParamInput, , lngCDID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "UpdateHardware", _
         "up_parmupd_hardware failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
  
   'Set the return code
   UpdateHardware = 0
   
   'Exit function
   Exit Function
   
UpdateHardware_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   UpdateHardware = 1
End Function
Public Function UpdateSoftware( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngSoftwareID As Long, _
   ByVal strSoftware As String, _
   ByVal lngCategoryID As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo UpdateSoftware_Err
   
   'Declare local variables and objects
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_software"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("SoftwareID", _
      adInteger, adParamInput, , lngSoftwareID)
   objCmd.Parameters.Append objCmd.CreateParameter("Software", _
      adVarChar, adParamInput, 30, strSoftware)
   objCmd.Parameters.Append objCmd.CreateParameter("CategoryID", _
      adInteger, adParamInput, , lngCategoryID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "UpdateSoftware", _
         "up_parmupd_software failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
  
   'Set the return code
   UpdateSoftware = 0
   
   'Exit function
   Exit Function
   
UpdateSoftware_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Return to the caller with a RC of 1
   UpdateSoftware = 1
End Function
Public Function UpdateSystemAssignment( _
   ByRef objConn As ADODB.Connection, _
   ByVal lngAssignmentID As Long, _
   ByVal lngEmployeeID As Long, _
   ByVal lngSystemID As Long, _
   ByRef lngSoftwareID() As Long, _
   ByRef strMessage As String) As Long

   'Setup error handling
   On Error GoTo UpdateSystemAssignment_Err
   
   'Declare local variables and objects
   Dim intIndex As Integer
   Dim objCmd As ADODB.Command
   
   'Set a reference to the command object
   Set objCmd = New ADODB.Command
   
   'Set the command object properties
   Set objCmd.ActiveConnection = objConn
   objCmd.CommandText = "up_parmupd_system_assignment"
   objCmd.CommandType = adCmdStoredProc
   
   'Create and append the parameters to the parameters collection
   objCmd.Parameters.Append objCmd.CreateParameter("RC", _
      adInteger, adParamReturnValue)
   objCmd.Parameters.Append objCmd.CreateParameter("SystemID", _
      adInteger, adParamInput, , lngAssignmentID)
   objCmd.Parameters.Append objCmd.CreateParameter("Employee", _
      adInteger, adParamInput, , lngEmployeeID)
   objCmd.Parameters.Append objCmd.CreateParameter("System", _
      adInteger, adParamInput, , lngSystemID)
   
   'Execute the command object
   objCmd.Execute
   
   'Check the return value from the stored procedure
   If objCmd.Parameters("RC") <> 0 Then
      Err.Raise 513 + vbObjectError, "UpdateSystemAssignment", _
         "up_parmupd_system_assignment failed"
   End If
   
   'Remove references to objects
   Set objCmd = Nothing
   
   'Process all software titles selected
   For intIndex = LBound(lngSoftwareID) To UBound(lngSoftwareID)
   
      'Set a reference to the command object
      Set objCmd = New ADODB.Command
      
      'Set the command object properties
      Set objCmd.ActiveConnection = objConn
      objCmd.CommandText = "up_parmins_system_software"
      objCmd.CommandType = adCmdStoredProc
      
      'Create and append the parameters to the parameters collection
      objCmd.Parameters.Append objCmd.CreateParameter("RC", _
         adInteger, adParamReturnValue)
      objCmd.Parameters.Append objCmd.CreateParameter("AssignedSystem", _
         adInteger, adParamInput, , lngAssignmentID)
      objCmd.Parameters.Append objCmd.CreateParameter("Software", _
         adInteger, adParamInput, , lngSoftwareID(intIndex))
      
      'Execute the command object
      objCmd.Execute
      
      'Check the return value from the stored procedure
      If objCmd.Parameters("RC") <> 0 Then
         Err.Raise 513 + vbObjectError, "UpdateSystemAssignment", _
            "up_parmins_system_software failed"
      End If
      
      'Remove references to objects
      Set objCmd = Nothing
   
   Next
  
   'Set the return code
   UpdateSystemAssignment = 0
   
   'Exit function
   Exit Function
   
UpdateSystemAssignment_Err:
   'Enumerate ADO errors
   Dim objErrors As New clsInsertData
   strMessage = objErrors.EnumerateErrors(objConn)
   Set objErrors = Nothing
   'Append any VB errors
   strMessage = strMessage & Err.Number & " : " & Err.Description
   'Clean up database
   objConn.Execute "up_parmdel_failed_system_assignment"
   'Return to the caller with a RC of 1
   UpdateSystemAssignment = 1
End Function


