VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanPelajaranGuru 
   Caption         =   "From Laporan Pelajaran Guru"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian Diperkecil"
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7890
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   6030
         TabIndex        =   7
         Top             =   585
         Width           =   1410
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   270
         Width           =   1410
      End
      Begin VB.ComboBox cmbSemester 
         Height          =   315
         Left            =   1485
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   630
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Semester"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   465
      Left            =   6660
      TabIndex        =   0
      Top             =   7155
      Width           =   1230
   End
   Begin MSComctlLib.ListView lstGuru 
      Height          =   5820
      Left            =   0
      TabIndex        =   1
      Top             =   1170
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   10266
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
Attribute VB_Name = "frmLaporanPelajaranGuru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAmbilKelas_Click()
    frmDataKelasLaporanRekap.Show , Me
End Sub

Private Sub cmdCari_Click()
    ListJadwal
End Sub

Private Sub cmdKeluar_Click()
    Visible = False
End Sub

Private Sub Command1_Click()
    frmDataGuruLaporanRekap.Show , Me
End Sub

Private Sub Form_Load()
    Clear Me

    Me.cmbSemester.Clear
    Me.cmbTahun.Clear

    Me.cmbSemester.Text = "--"
    Me.cmbTahun.Text = "--"

    Me.cmbSemester.AddItem "--"
    Me.cmbTahun.AddItem "--"


    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select * from semester where deleted = 0")

    If Not rs.EOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            Me.cmbSemester.AddItem rs.Fields(1)
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)

    Set rs = openRecordset("select distinct(tahun) from jadwal where deleted = 0")

    If Not rs.EOF Then
        rs.MoveFirst

        Do While Not rs.EOF
            Me.cmbTahun.AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)

    ListJadwal

End Sub

Sub ListJadwal()
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

        Dim rs As ADODB.Recordset
        Dim rssemester As ADODB.Recordset
        Dim query As String

        ' query = "select guru.*, kelas.nama as namakelas, semester.nama as namasemester, guru.nama as namaguru, pelajaran.nama as namapelajaran  from jadwal inner join kelas on kelas.id = jadwal.idkelas inner join semester on semester.id = jadwal.semester inner join guru on guru.id = jadwal.idguru inner join pelajaran on pelajaran.id = jadwal.idpelajaran where guru.deleted = 0 and semester.deleted = 0 and jadwal.deleted = 0 and kelas.deleted = 0 "
        query = "select guru.id ,guru.nama, guru.jeniskelamin, guru.nip, guru.status,  guru.agama, guru.tempat, guru.tanggallahir  from jadwal inner join semester on semester.id = jadwal.semester inner join guru on guru.id = jadwal.idguru where guru.deleted = 0 and semester.deleted = 0 and jadwal.deleted = 0 "

        If Not cmbTahun.Text = "--" Then
            query = query & " and tahun = '" & Trim$(cmbTahun.Text) & "'"
        End If


        If Not cmbSemester.Text = "--" Then
            Set rssemester = openRecordset("select * from semester where deleted = 0 and nama = '" & cmbSemester.Text & "'")

            If Not rssemester.EOF Then
                query = query & " and semester = '" & Trim$(rssemester.Fields(0)) & "'"
            End If

            Call closeRecordset(rssemester)
        End If

        query = query & " group by guru.id ,guru.nama, guru.jeniskelamin, guru.nip, guru.status,  guru.agama, guru.tempat, guru.tanggallahir"

        Set rs = openRecordset(query)

        .ListItems.Clear

        Dim i As Integer
        i = 1

        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Dim List As ListItem
                Set List = .ListItems.Add(, , i)

                fillListView List, rs, "id"
                fillListView List, rs, "Nama"
                fillListView List, rs, "jeniskelamin"
                fillListView List, rs, "nip"
                fillListView List, rs, "status"
                fillListView List, rs, "agama"
                fillListView List, rs, "tempat"
                fillListView List, rs, "tanggallahir"

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)

    End With
End Sub

Private Sub lstJadwal_DblClick()

End Sub

Private Sub lstGuru_DblClick()
    On Error Resume Next
    If Me.lstGuru.ListItems.Count = 0 Then Exit Sub

    Dim tahun As String
    Dim semester As String
    Dim rssemester As ADODB.Recordset

    If Not cmbTahun.Text = "--" Then
        tahun = Trim$(cmbTahun.Text)
    Else
        MsgBox "Tahun Harus dipilih"
        Exit Sub
    End If


    If Not cmbSemester.Text = "--" Then
        Set rssemester = openRecordset("select * from semester where deleted = 0 and nama = '" & cmbSemester.Text & "'")

        If Not rssemester.EOF Then
            semester = Trim$(rssemester.Fields(0))
        End If

        Call closeRecordset(rssemester)

    Else
        MsgBox "Semester harus dipilih"
        Exit Sub
    End If

    With DataEnvironmentGuru.rsCommandJadwalGuru_Grouping
        If Not .State = 0 Then .Close
        DataEnvironmentGuru.CommandJadwalGuru_Grouping lstGuru.SelectedItem.SubItems(1), semester, tahun
        ReportJadwalGuru.Show , Me
    End With
End Sub
