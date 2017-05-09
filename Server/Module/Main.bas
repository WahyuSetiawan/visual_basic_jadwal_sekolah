Attribute VB_Name = "ModMain"
Public Sub Main()
    If konekToServer Then
        frmLogin.Show
    Else
        MsgBox "Connection Problem, please Setting your connection"
        frmConnectionProblemdatabase.Show
    End If
End Sub
