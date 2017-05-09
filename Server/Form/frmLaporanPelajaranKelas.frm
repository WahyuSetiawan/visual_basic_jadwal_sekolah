VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLaporanPelajaranKelas 
   Caption         =   "Form Laporan Pelajaran Kelas"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian Diperkecil"
      Height          =   1230
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   7890
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   5895
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
      Left            =   6705
      TabIndex        =   0
      Top             =   7245
      Width           =   1230
   End
   Begin MSComctlLib.ListView lstKelas 
      Height          =   5685
      Left            =   45
      TabIndex        =   1
      Top             =   1395
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   10028
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
Attribute VB_Name = "frmLaporanPelajaranKelas"
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
    With Me.lstKelas

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Nama", 2000

        Dim rs As ADODB.Recordset
        Dim rssemester As ADODB.Recordset
        Dim query As String

        query = "select kelas.id, kelas.nama as namakelas  from jadwal inner join kelas on kelas.id = jadwal.idkelas inner join semester on semester.id = jadwal.semester inner join guru on guru.id = jadwal.idguru inner join pelajaran on pelajaran.id = jadwal.idpelajaran where guru.deleted = 0 and semester.deleted = 0 and jadwal.deleted = 0 and kelas.deleted = 0 "

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

        query = query & " group by kelas.id, kelas.nama"

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
                fillListView List, rs, "namakelas"

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)

    End With
End Sub


Private Sub lstKelas_DblClick()
    On Error Resume Next

    If Me.lstKelas.ListItems.Count = 0 Then Exit Sub

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
    With DataEnvironmentGuru.rsCommandJadwalKelas_Grouping
        If Not .State = 0 Then .Close
        DataEnvironmentGuru.CommandJadwalKelas_Grouping lstKelas.SelectedItem.SubItems(1), semester, tahun
        ReportJadwalKelas.Show , Me
    End With
End Sub
