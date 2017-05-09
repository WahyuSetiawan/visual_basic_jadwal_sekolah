Attribute VB_Name = "Connection"
Public Conn As New adodb.Connection
Public RS As New adodb.Recordset
Public RS1 As New adodb.Recordset
Public RS2 As New adodb.Recordset
Public RS3 As New adodb.Recordset
Public RS4 As New adodb.Recordset
Public Strconn As String
Public strsql As String
Public no As Integer

Sub KONEKSI()
'Strconn = "Provider=SQLOLEDB.1;Password=abcd1234;Persist Security Info=True;User ID=sa;Initial Catalog=aplikasipresensi;Data Source=FOD-PC"
    OpenDatabaseSQLServer "FOD-PC", "sa", "abcd1234", "aplikasipresensi"
End Sub

Sub OpenDatabaseSQLServer(hostname As String, username As String, password As String, database As String)
    Strconn = "Provider=SQLOLEDB.1;Password=" + password + ";Persist Security Info=True;User ID=" + username + ";Initial Catalog=" + database + ";Data Source=" + hostname
    Conn.CursorLocation = adUseClient
    If Conn.State = adStateClosed Then
        Conn.Open Strconn
    End If
End Sub


