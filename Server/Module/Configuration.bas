Attribute VB_Name = "Configuration"
Public Const pathFileConfiguration As String = "\conf.dat"

Sub createAndSaveFile(path As String, isi() As String)
    Dim intFile As Integer
    Dim i As Integer

    intFile = FreeFile

    Open App.path & path For Output As #intFile

    For i = LBound(isi) To UBound(isi)
        Print #intFile, isi(i)
    Next i

    Close #intFile
End Sub


Function loadDataFromFile(path As String) As String()
    Dim DataTMP() As String
    Dim intFile As Integer
    Dim i As Integer
    Dim tmp As String

    i = 0

    intFile = FreeFile

    If Not Dir$(App.path & path) = "" Then
        Open App.path & path For Input As #intFile

        Do Until EOF(intFile)
            ReDim Preserve DataTMP(i)

            Input #intFile, tmp

            DataTMP(i) = tmp
            i = i + 1
        Loop

        loadDataFromFile = DataTMP
    End If

    Close #intFile
End Function

Function getDataFromArray(data() As String, parameter As String) As String
    Dim i As Integer
    Dim DataTMP() As String

    For i = LBound(data) To UBound(data)
        DataTMP = Split(data(i), ":")

        If DataTMP(0) = parameter Then
            getDataFromArray = DataTMP(1)

            Close #intFile
            Exit Function
        End If
    Next i

    getDataFromArray = ""
    Close #intFile
End Function

Function getDataFromFile(path As String, parameter As String) As String
    Dim dataFile() As String
    Dim i As Integer

    dataFile = loadDataFromFile(path)

    getDataFromFile = getDataFromArray(dataFile, parameter)
End Function


