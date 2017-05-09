VERSION 5.00
Begin VB.Form frmLaporanPelajaranDanGruu 
   Caption         =   "From Pelajaran dan guru"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   510
      Left            =   4635
      TabIndex        =   6
      Top             =   1125
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetakSemua 
      Caption         =   "Cetak Semua"
      Height          =   510
      Left            =   3375
      TabIndex        =   5
      Top             =   1125
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian Diperkecil"
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5820
      Begin VB.ComboBox cmbTahun 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   270
         Width           =   1410
      End
      Begin VB.ComboBox cmbSemester 
         Height          =   315
         Left            =   1485
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   630
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   465
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
   End
End
Attribute VB_Name = "frmLaporanPelajaranDanGruu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCetakSemua_Click()
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

    With DataEnvironmentGuru.rsCommandPelajaranDanGuru_Grouping
        If Not .State = 0 Then .Close
        DataEnvironmentGuru.CommandPelajaranDanGuru_Grouping semester, tahun
        ReportPelajaranuGuru.Show , Me
    End With
End Sub

Private Sub cmdKeluar_Click()
    Visible = False
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
End Sub
