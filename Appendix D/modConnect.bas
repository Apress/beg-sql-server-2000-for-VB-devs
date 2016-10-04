Attribute VB_Name = "modConnect"
Option Explicit

'Declare public objects
Public g_objConn As ADODB.Connection
Public g_objError As ADODB.Error
Public g_objRS As ADODB.Recordset



Public Function EstablishConnection( _
   ByVal blnDSN As Boolean, _
   ByVal strDSN As String, _
   ByVal blnWindowsAuthentication As Boolean, _
   ByVal strLogin As String, _
   ByVal strPassword As String, _
   ByVal blnConnectString As Boolean) As Boolean

   'Setup error handling
   On Error GoTo EstablishConnection_Error
   
   'Declare local variables
   Dim strConnectString As String
   
   'Set a reference to the ADO Connection object
   Set g_objConn = New ADODB.Connection
   
   If Not blnDSN And blnWindowsAuthentication Then
   
    g_objConn.Provider = "SQLOLEDB"
    g_objConn.Properties("Data Source").Value = "WSTRAVEL\SQL2000"
    g_objConn.Properties("Initial Catalog").Value = "Hardware Tracking"
    g_objConn.Properties("Integrated Security").Value = "SSPI"
    
    g_objConn.Open
   
   Else
   
      'Build the DSN or DSN-Less part of the connect string
      If blnDSN Then
         strConnectString = "DSN=" & strDSN & ";"
      Else
         strConnectString = "Provider=SQLOLEDB;" & _
                           "Data Source=WSTRAVEL\SQL2000;" & _
                           "Initial Catalog=Hardware Tracking;"
      End If
      
      'Add the User ID and Password to the connect string
      If Not blnWindowsAuthentication And blnConnectString Then
         strConnectString = strConnectString & _
            "User ID=" & strLogin & ";" & _
            "Password=" & strPassword & ";"
      End If
      
      'Open the Connection object
      If blnWindowsAuthentication Or blnConnectString Then
         g_objConn.Open strConnectString
      Else
         g_objConn.Open strConnectString, strLogin, strPassword
      End If
      
   End If
   
   'Check connection state
   If g_objConn.State <> adStateOpen Then
      EstablishConnection = False
   Else
      EstablishConnection = True
   End If
   
   'Exit function
   Exit Function
   
EstablishConnection_Error:
   'Connection failed, display error messages
   Dim strError
   For Each g_objError In g_objConn.Errors
      strError = strError & g_objError.Number & " : " & _
         g_objError.Description & vbCrLf & vbCrLf
   Next
   MsgBox strError, vbCritical + vbOKOnly, "Login Error"
End Function

Public Sub TerminateConnection()
   On Error Resume Next
   g_objRS.Close
   Set g_objRS = Nothing
   g_objConn.Close
   Set g_objConn = Nothing
   Set g_objError = Nothing
End Sub


