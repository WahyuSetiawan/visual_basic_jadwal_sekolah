VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanDataGuru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Data Guru"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   510
      Left            =   8415
      TabIndex        =   6
      Top             =   7155
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetakSemua 
      Caption         =   "Cetak Semua"
      Height          =   510
      Left            =   7155
      TabIndex        =   5
      Top             =   7155
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mencari Data Guru"
      Height          =   690
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9510
      Begin VB.CommandButton Command1 
         Caption         =   "Cari"
         Height          =   420
         Left            =   8235
         TabIndex        =   2
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1305
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   225
         Width           =   6765
      End
      Begin VB.Label Label1 
         Caption         =   "Data Guru"
         Height          =   375
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1410
      End
   End
   Begin MSComctlLib.ListView lstGuru 
      Height          =   6360
      Left            =   90
      TabIndex        =   4
      Top             =   765
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
Attribute VB_Name = "FrmLaporanDataGuru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub fillDataAnggota()
    With Me.lstGuru

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Nama", 2000
        .ColumnHeaders.Add , , "Jenis Kelamin", 500
        .ColumnHeaders.Add , , "Nip", 1000
        .ColumnHeaders.Add , , "Status", 1000
        .ColumnHeaders.Add , , "Agama", 800
        .ColumnHeaders.Add , , "Tempat", 1500
        .ColumnHeaders.Add , , "Tanggal Lahir", 1000

        Dim rsDB As ADODB.Recordset

        Set rsDB = openRecordset("select * from Guru where id like '%" + Trim$(parameter) + "%' and deleted = 0")

        .ListItems.Clear

        Dim i As Integer
        i = 1

        If Not rsDB.EOF Then
            rsDB.MoveFirst
            Do While Not rsDB.EOF
                Dim List As ListItem
                Set List = .ListItems.Add(, , i)

                fillListView List, rsDB, "id"
                fillListView List, rsDB, "Nama"
                fillListView List, rsDB, "jeniskelamin"
                fillListView List, rsDB, "nip"
                fillListView List, rsDB, "status"
                fillListView List, rsDB, "agama"
                fillListView List, rsDB, "tempat"
                fillListView List, rsDB, "tanggallahir"

                i = i + 1
                rsDB.MoveNext
            Loop
        End If

        Call closeRecordset(rsDB)
    End With
End Sub

Private Sub cmdCetakSemua_Click()
    With DataEnvironmentGuru.rsCommandDataGuru
        If Not .State = 0 Then .Close
        DataEnvironmentGuru.CommandDataGuru
        Printer.Orientation = vbPRORLandscape
        ReportDataGuru.Orientation = rptOrientLandscape
        ReportDataGuru.Show , Me
    End With
End Sub

Private Sub cmdKeluar_Click()
    Visible = False
End Sub

Private Sub Command1_Click()
    fillDataAnggota
End Sub

Private Sub Form_Load()
    Clear Me
    fillDataAnggota
End Sub


Private Sub lstGuru_DblClick()
    If lstGuru.ListItems.Count = 0 Then Exit Sub
    With DataEnvironmentGuru.rsCommandDataGuru
        If Not .State = 0 Then .Close
        .Open
        .Requery
        .Filter = " id = " & lstGuru.SelectedItem.SubItems(1)
        ReportDataGuruISingle.Show , Me
    End With
End Sub
