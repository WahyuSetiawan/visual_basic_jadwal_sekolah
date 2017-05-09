VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanRekapPresensi 
   Caption         =   "Form Laporan Rekap Jadwal"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   465
      Left            =   6705
      TabIndex        =   15
      Top             =   7200
      Width           =   1230
   End
   Begin MSComctlLib.ListView lstJadwal 
      Height          =   5100
      Left            =   45
      TabIndex        =   5
      Top             =   1935
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   8996
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian Diperkecil"
      Height          =   1860
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7890
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   6075
         TabIndex        =   14
         Top             =   1305
         Width           =   1410
      End
      Begin VB.CheckBox CheckGuru 
         Caption         =   "Guru"
         Height          =   510
         Left            =   135
         TabIndex        =   13
         Top             =   1260
         Width           =   1230
      End
      Begin VB.CheckBox CheckKelas 
         Caption         =   "Kelas"
         Height          =   510
         Left            =   135
         TabIndex        =   12
         Top             =   855
         Width           =   1230
      End
      Begin VB.TextBox txtGuru 
         Height          =   330
         Left            =   1485
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1350
         Width           =   3840
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "..."
         Height          =   330
         Left            =   5310
         TabIndex        =   8
         Top             =   1350
         Width           =   285
      End
      Begin VB.CommandButton cmdAmbilKelas 
         Caption         =   "..."
         Height          =   330
         Left            =   5310
         TabIndex        =   7
         Top             =   990
         Width           =   285
      End
      Begin VB.TextBox txtKelas 
         Height          =   330
         Left            =   1485
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   990
         Width           =   3840
      End
      Begin VB.ComboBox cmbSemester 
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   630
         Width           =   4110
      End
      Begin VB.ComboBox cmbTahun 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label lblGuru 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         Height          =   195
         Left            =   6255
         TabIndex        =   10
         Top             =   945
         Width           =   480
      End
      Begin VB.Label lblKelas 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         Height          =   195
         Left            =   6255
         TabIndex        =   9
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Semester"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmLaporanRekapPresensi"
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

    lblGuru.Caption = ""
    lblKelas.Caption = ""

    Me.cmbSemester.Clear
    Me.cmbTahun.Clear

    Me.txtGuru.Text = "--"
    Me.cmbSemester.Text = "--"
    Me.cmbTahun.Text = "--"
    Me.txtKelas.Text = "--"

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
    With Me.lstJadwal

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Hari", 500
        .ColumnHeaders.Add , , "Jadwal Mulai", 1500
        .ColumnHeaders.Add , , "Jadwal Selesai", 1500
        .ColumnHeaders.Add , , "Id Pelajaran", 1000
        .ColumnHeaders.Add , , "Id Kelas", 1000
        .ColumnHeaders.Add , , "Id Guru", 1000

        Dim rs As ADODB.Recordset
        Dim rssemester As ADODB.Recordset
        Dim query As String

        query = "select jadwal.*, " & _
                "(CASE WHEN hari=1 then 'Minggu'" & _
                "WHEN hari=2 THEN 'Senin'" & _
                "WHEN hari=3 THEN 'Selasa'" & _
                "WHEN hari=4 THEN 'Rabu'" & _
                "WHEN hari=5 THEN 'Kamis'" & _
                "WHEN hari=6 THEN 'Jumat' ELSE 'Sabtu' END ) as namahari, " & _
                "kelas.nama as namakelas, semester.nama as namasemester, guru.nama as namaguru, pelajaran.nama as namapelajaran  from jadwal inner join kelas on kelas.id = jadwal.idkelas inner join semester on semester.id = jadwal.semester inner join guru on guru.id = jadwal.idguru inner join pelajaran on pelajaran.id = jadwal.idpelajaran where guru.deleted = 0 and semester.deleted = 0 and jadwal.deleted = 0 and kelas.deleted = 0 "

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


        If (Not lblGuru.Caption = "--") And CheckGuru.value Then
            query = query & " and idguru = '" & Trim$(lblGuru.Caption) & "'"
        End If


        If (Not lblKelas.Caption = "") And CheckKelas.value Then
            query = query & " and idkelas = '" & Trim$(lblKelas.Caption) & "'"
        End If

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
                fillListView List, rs, "namahari"
                fillListView List, rs, "waktumulai"
                fillListView List, rs, "waktuselesai"
                fillListView List, rs, "namapelajaran"
                fillListView List, rs, "namakelas"
                fillListView List, rs, "namaguru"

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)

    End With
End Sub

Private Sub lstJadwal_DblClick()
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("Select " & _
                           "jadwal.idguru, " & _
                           "rekapjadwal.keterangan, " & _
                           "jadwal.waktumulai, " & _
                           "jadwal.waktuselesai, " & _
                           "semester.nama as namasemester," & _
                           "jadwal.tahun," & _
                           "convert(varchar(10), tanggal, 105) as tanggal, " & _
                           "(CASE WHEN hari=1 then 'Minggu' " & _
                           "WHEN hari=2 THEN 'Senin' " & _
                           "WHEN hari=3 THEN 'Selasa' " & _
                           "WHEN hari=4 THEN 'Rabu' " & _
                           "WHEN hari=5 THEN 'Kamis' " & _
                           "WHEN hari=6 THEN 'Jumat' ELSE 'Sabtu' END ) as namahari, " & _
                           "guru.nama as namaguru, " & _
                           "guru.nip, " & _
                           "pelajaran.nama as namapelajaran, " & _
                           "kelas.nama as namakelas " & _
                           "From jadwal " & _
                           "inner join guru on guru.id = jadwal.idguru " & _
                           "inner join pelajaran on pelajaran.id = jadwal.idpelajaran " & _
                           "inner join kelas on kelas.id = jadwal.idkelas " & _
                           "inner join semester on semester.id = jadwal.semester " & _
                           "inner join rekapjadwal on rekapjadwal.idjadwal = jadwal.id " & _
                           "where jadwal.deleted = 0  and guru.deleted = 0 and pelajaran.deleted = 0 and semester.deleted = 0 and idjadwal = " & lstJadwal.SelectedItem.SubItems(1))

    If rs.EOF Then
        MsgBox "tidak terdapat data yang akan ditampilkan"
        Exit Sub
    End If

    If Me.lstJadwal.ListItems.Count = 0 Then Exit Sub
    With DataEnvironmentGuru.rsCommandRekapJadwal_Grouping
        If Not .State = 0 Then .Close
        DataEnvironmentGuru.CommandRekapJadwal_Grouping Me.lstJadwal.SelectedItem.SubItems(1)
        ReportRekapPresensi.Show , Me
    End With
End Sub
