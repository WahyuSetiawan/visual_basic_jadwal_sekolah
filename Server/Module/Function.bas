Attribute VB_Name = "Function"
Public Type ArrayList
    nama As String
End Type

Public Function CloseAllForms() As Integer
    answer = MsgBox("Apakah anda yakin untuk menutup aplikasi ini ?", vbOKCancel)
    If answer = vbOK Then
        CloseAllForms = vbOK
        Dim Frm As Form
        For Each Frm In Forms
            Unload Frm
            Set Frm = Nothing
        Next
        Exit Function
    End If
    CloseAllForms = vbCancel
End Function

Public Function Clean(Text)
    On Error Resume Next
    ' Dim Chars As ?????????
    Chars = Array("txt_", "_", "cmd_", "dtp_")
    For Each Replaced In Chars
        Text = Replace(Text, Replaced, "")
    Next
    Clean = CStr(Text)
End Function

Public Function NotNull(Control As Control) As Boolean
    If TypeOf Control Is TextBox Then
        If Control.Text = "" Then
            MsgBox Clean(Control.Name) & " Jangan Kosong !"
            Control.SetFocus
            NotNull = False
            Exit Function
        End If
    End If

    If TypeOf Control Is ComboBox Then
        If Control.Text = "" Then
            MsgBox Clean(Control.Name) & " belum dipilih !"
            Control.SetFocus
            NotNull = False
            Exit Function
        End If
    End If

    NotNull = True
End Function

Public Sub CheckFormat(Control As Control, Format As String)

End Sub

Public Function CheckNull(Form As Form) As Boolean
    Dim X As Control
    For Each X In Form
        If TypeOf X Is TextBox Then
            If X.Text = "" Then
                MsgBox Replace(Replace(X.Name, "txt_", ""), "_", " ") & " Masih Kosong"
                If X.Enabled = True Then
                    X.SetFocus
                End If
                CheckNull = False
                Exit Function
            End If
        End If
    Next
    For Each X In Form
        If TypeOf X Is ComboBox Then
            If X.Text = "" Then
                MsgBox Replace(Replace(X.Name, "txt_", ""), "_", " ") & " Masih Kosong"
                If X.Enabled = True Then
                    X.SetFocus
                End If
                CheckNull = False
                Exit Function
            End If
        End If
    Next

    CheckNull = True
End Function

Public Sub Clear(Form As Form)
    Dim X As Control
    For Each X In Form
        If TypeOf X Is TextBox Then X.Text = ""
    Next

    For Each X In Form
        If TypeOf X Is ComboBox Then X.Text = ""
    Next

    For Each X In Form
        If TypeOf X Is CheckBox Then X.value = 0
    Next
End Sub

Public Function DelHandle(Text As String) As Boolean
    If Text = "" Then
        DelHandle = False
        Exit Function
    End If

    If vbNo = MsgBox(delMessage & Text, vbYesNo, "Konfirmasi") Then
        DelHandle = False
    End If

    DelHandle = True
End Function

Public Function IsNumber(Control As Control) As Boolean
    IsNumber = False

    If TypeOf Control Is TextBox Then
        If Control.Text = DisableText Then Exit Function
        If Control.Enabled Then
            If Control.Text = "" Then Control.Text = 0
            If Not IsNumeric(Control.Text) Then
                MsgBox Control.Name & numMessage
                Control.SetFocus
                Control.Text = 0
                Exit Function
            End If
        End If
    End If

    If TypeOf Control Is ComboBox Then
        If Control.Text = DisableText Then Exit Function
        If Control.Enabled Then
            If Control.Text = "" Then Control.Text = 0
            If Not IsNumeric(Control.Text) Then
                MsgBox Control.Name & numMessage
                Control.SetFocus
                Control.Text = 0
                Exit Function
            End If
        End If
    End If

    IsNumber = True
End Function

Public Function NumberORVarchar(Text As String) As String
    Text = Trim$(Text)
    If Not IsNumeric(Text) Then
        NumberORVarchar = "'" & Text & "'"
    Else
        NumberORVarchar = "" & Text & ""
    End If
End Function

Public Function FormatDateSQL(datestring As String) As String
    datestring = Format$(datestring, "yyyy-mm-dd")
    FormatDateSQL = datestring
End Function

Public Function FormatTimeSQL(TimeString As String) As String
    TimeString = Format$(TimeString, "h:mm:ss")
    FormatTimeSQL = TimeString
End Function


Public Sub autoAddItem(Combo As ComboBox, rs As Recordset, Field As String)
    On Error GoTo Error:
    If rs.RecordCount = 0 Then Exit Sub

    rs.MoveFirst
    Do
        Combo.AddItem rs.Fields(Field)
        rs.MoveNext
    Loop Until rs.EOF

    Set rs = Nothing
    Exit Sub
Error:
    MsgBox messageErrorCombo
    Set rs = Nothing
End Sub

Public Sub checkUpDown(Control As Control, Down As Integer, UP As Integer)
    If Control.Text = "" Then Exit Sub
    If Not IsNumber(Control) Then Exit Sub

    If (CInt(Control.Text) < Down) Then
        Control.Text = Down
        MsgBox "range " & Down & " hingga " & UP
    End If
    If (CInt(Control.Text) > UP) Then
        Control.Text = UP
        MsgBox "range " & Down & " hingga " & UP
    End If
End Sub

Public Sub DisableTextBox(myForm As Form)
    Dim X As Control
    For Each X In myForm
        If TypeOf X Is TextBox Then
            X.Text = DisableText
            X.Enabled = False
        End If
    Next

    For Each X In myForm
        If TypeOf X Is ComboBox Then
            X.Text = DisableText
            X.Enabled = False
        End If
    Next

    For Each X In myForm
        If TypeOf X Is DTPicker Then
            X.Enabled = False
        End If
    Next

    For Each X In myForm
        If TypeOf X Is OptionButton Then
            X.Enabled = False
        End If
    Next
End Sub

Public Sub EnableTextBox(myForm As Form)
    Dim X As Control
    For Each X In myForm
        If TypeOf X Is TextBox Then X.Enabled = True
    Next

    For Each X In myForm
        If TypeOf X Is ComboBox Then X.Enabled = True
    Next

    For Each X In myForm
        If TypeOf X Is DTPicker Then X.Enabled = True
    Next

    For Each X In myForm
        If TypeOf X Is OptionButton Then X.Enabled = True
    Next
End Sub

Public Sub IncrementValue(Control As Control, Last As Integer, value As Integer)
    If Control.Text = DisableText Then Exit Sub
    If Control.Text = "" Then Control.Text = "0"
    If CInt(Control.Text) < Last Then
        Control.Text = CInt(Control.Text) + value
    End If
End Sub

Public Sub DecrementValue(Control As Control, Last As Integer, value As Integer)
    If Control.Text = DisableText Then Exit Sub
    If Control.Text = "" Then Control.Text = "0"
    If CInt(Control.Text) > Last Then
        Control.Text = CInt(Control.Text) - value
    End If
End Sub

Public Function notAllowChange() As Boolean
    notAllowChange = True
End Function

Public Function generateDateFormMonthAndYear(month As String, year As String) As String
    generateDateFormMonthAndYear = Format$("30/" & month & "/" & year, "yyyy-mm-dd")
End Function

Public Function getMonth(datestring As String) As Integer
    getMonth = CInt(Format$(datestring, "mm"))
End Function

Public Function getYear(datestring As String) As Integer
    getYear = CInt(Format$(datestring, "yyyy"))
End Function

Public Function getDateFormatSql(s As String) As String
    getDateFormatSql = Format$(s, "yyyy-mm-dd")
End Function

Public Function setDtpickerFromSql() As String

End Function

Public Function fillListView(List As ListItem, rs As Recordset, parameter As String)
    If IsNull(rs.Fields(parameter).value) = True Then
        List.ListSubItems.Add , , ""
    Else
        List.ListSubItems.Add , , rs.Fields(parameter)
    End If
End Function

Public Function HariDariUrutanTanggal(urutan As Integer) As String
    Select Case urutan
        Case 1
            HariDariUrutanTanggal = "Minggu"
        Case 2
            HariDariUrutanTanggal = "Senin"
        Case 3
            HariDariUrutanTanggal = "Selasa"
        Case 4
            HariDariUrutanTanggal = "Rabu"
        Case 5
            HariDariUrutanTanggal = "Kamis"
        Case 6
            HariDariUrutanTanggal = "Jumat"
        Case 7
            HariDariUrutanTanggal = "Sabtu"
    End Select
End Function

Public Function UrutanDariHari(hari As String) As Integer
    Select Case hari
        Case "Senin"
            UrutanDariHari = 2
        Case "Selasa"
            UrutanDariHari = 3
        Case "Rabu"
            UrutanDariHari = 4
        Case "Kamis"
            UrutanDariHari = 5
        Case "Jumat"
            UrutanDariHari = 6
        Case "Sabtu"
            UrutanDariHari = 7
    End Select
End Function

Public Function secondToDate(waktu As Double) As String
    Dim seconds As Double
    Dim minute As Double
    Dim hours As Double
    
    minute = waktu \ 60
    hours = minute \ 60
    seconds = waktu - (minute * 60)
    minute = minute - (hours * 60)

    secondToDate = Right$("0" & CStr(hours), 2) & ":" & Right$("0" & CStr(minute), 2) & ":" & Right$("0" & CStr(seconds), 2)
End Function

Public Function dateToSeconds(datea As Date) As Double
    dateToSeconds = CDate(CStr(Format(datea, "HH:mm:ss"))) * 60 * 60 * 24
End Function


