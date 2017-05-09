Attribute VB_Name = "Connection"
Public Conn As New ADODB.Connection
Public Strconn As String
Public strSql As String
Public no As Integer

Sub KONEKSI()
    'Strconn = "Provider=SQLOLEDB.1;Password=abcd1234;Persist Security Info=True;User ID=sa;Initial Catalog=aplikasipresensi;Data Source=FOD-PC"
    Dim dataFile() As String
    Dim i As Integer

    Dim file As String

    file = "\connectionconf.dat"

    Dim host As String
    Dim username As String
    Dim password As String
    Dim database As String

    If Dir$(App.path & file) = "" Then
        Dim data() As String

        ReDim Preserve data(0)
        data(0) = "host:localhost"
        ReDim Preserve data(1)
        data(1) = "username:sa"
        ReDim Preserve data(2)
        data(2) = "password:123"
        ReDim Preserve data(3)
        data(3) = "database:databaseaplikasi"

        createAndSaveFile file, data
    End If

    dataFile = loadDataFromFile(file)

    host = getDataFromArray(dataFile, "host")
    username = getDataFromArray(dataFile, "username")
    password = getDataFromArray(dataFile, "password")
    database = getDataFromArray(dataFile, "database")

    OpenDatabaseSQLServer host, username, password, database
End Sub

Sub OpenDatabaseSQLServer(hostname As String, username As String, password As String, database As String)
    Strconn = "Provider=SQLOLEDB.1;Password=" + password + ";Persist Security Info=True;User ID=" + username + ";Initial Catalog=" + database + ";Data Source=" + hostname
    Conn.CursorLocation = adUseClient
    If Conn.State = adStateClosed Then
        Conn.Open Strconn
    End If
End Sub

Public Function testKonekToServer(hostname As String, username As String, password As String, database As String) As Boolean
    On Error GoTo errHandle

    Dim a123 As New ADODB.Connection


    Strconn = "Provider=SQLOLEDB.1;Password=" + password + ";Persist Security Info=True;User ID=" + username + ";Initial Catalog=" + database + ";Data Source=" + hostname
    a123.CursorLocation = adUseClient
    If a123.State = adStateClosed Then
        a123.Open Strconn
    End If

    testKonekToServer = True

    Exit Function
errHandle:
    testKonekToServer = False
End Function

Public Function konekToServer() As Boolean
    Dim strCon As String

    On Error GoTo errHandle

    KONEKSI

    konekToServer = True

    Exit Function
errHandle:
    konekToServer = False
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next

    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then
            vRs.Close
            Set vRs = Nothing
        End If
    End If
End Sub

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue As ADODB.Recordset

    On Error GoTo errHandle

    Set rsDbGetValue = New ADODB.Recordset
    rsDbGetValue.Open query, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsDbGetValue.EOF Then
        If Not IsNull(rsDbGetValue(0).value) Then
            dbGetValue = rsDbGetValue(0).value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If

    Call closeRecordset(rsDbGetValue)

    Exit Function
errHandle:
    dbGetValue = defValue
End Function

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset

    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, Conn, adOpenDynamic, adLockOptimistic
    Set openRecordset = obj
End Function

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
    On Error Resume Next

    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function





