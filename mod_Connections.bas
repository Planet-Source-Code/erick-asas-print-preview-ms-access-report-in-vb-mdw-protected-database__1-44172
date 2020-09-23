Attribute VB_Name = "mod_Connections"
Option Explicit

Private Conn As ADODB.Connection

Public Function Establish_Connection_DSN(ByVal struid As String, ByVal strpwd As String) As ADODB.Connection
On Error GoTo ET
  Dim strConn As String

  'Connect to database
  Set Conn = New ADODB.Connection
  Conn.CursorLocation = adUseClient
  strConn = "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;dsn=" & "MSAccessReportDemo" & ";uid=" & _
            struid & ";pwd=" & _
            strpwd & ";database=" & "ESIS" & ";"
  'Connect
  Conn.Open strConn

  Set Establish_Connection_DSN = Conn

  Exit Function

ET:
  Err.Raise Err.Number, Err.Description

    Exit Function

End Function

Public Function Establish_Connection_DSNLess(ByVal struid As String, ByVal strpwd As String) As ADODB.Connection
On Error GoTo ET
  Dim strConn As String

  'Connect to database
  Set Conn = New ADODB.Connection
  Conn.CursorLocation = adUseClient
  Conn.CursorLocation = adUseClient
  'We'll use a dsn-less connection
  'Define DSN-less string
  strConn = "DRIVER=Microsoft Access Driver (*.mdb);" & _
    "DBQ=" & App.Path & "\ESIS.mdb;" & _
    "SystemDB=" & App.Path & "\ESIS.mdw;" & _
    "UID= " & struid & _
    ";PWD= " & strpwd
  'Connect
  Conn.Open strConn

  Set Establish_Connection_DSNLess = Conn

  Exit Function

ET:
  Err.Raise Err.Number, Err.Description

    Exit Function

End Function


