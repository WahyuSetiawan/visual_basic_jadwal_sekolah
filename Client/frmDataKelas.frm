VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataKelas 
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Mencari Data Pelajaran"
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9510
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1305
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   225
         Width           =   6765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cari"
         Height          =   420
         Left            =   8235
         TabIndex        =   1
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Data Pelajaran"
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1410
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6360
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   11218
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDataKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const delimiterBaris As String = "|"
Const delimiterKolom As String = "#"

Sub fillDataAnggota()
    With Me.ListView1

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Nama", 2000

        Dim i As Integer
        Dim dataSplit() As String
        Dim dataKolom() As String

        dataSplit = Split(getDataFromFile("\dataTMP.dat", "datakelas"), delimiterBaris)

        For i = LBound(dataSplit) To UBound(dataSplit)

            If Not dataSplit(i) = "" Then
                Dim List As ListItem
                Set List = .ListItems.Add(, , i + 1)

                dataKolom = Split(dataSplit(i), delimiterKolom)
                List.ListSubItems.Add , , CStr(dataKolom(0))
                List.ListSubItems.Add , , CStr(dataKolom(1))
            End If
        Next i
    End With
End Sub

Private Sub Command1_Click()
    fillDataAnggota
End Sub

Private Sub Form_Load()
    Clear Me
    fillDataAnggota
End Sub


Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub

    Load frmClientUtama
    frmClientUtama.idKelas = CInt(ListView1.SelectedItem.SubItems(1))
    frmClientUtama.txtKelas.Text = ListView1.SelectedItem.SubItems(1) + " - " + ListView1.SelectedItem.SubItems(2)
    frmClientUtama.Show
    Unload Me
End Sub
