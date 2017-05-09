VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListJadwal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSemester 
      Height          =   315
      Left            =   4770
      TabIndex        =   2
      Text            =   "cmbSemester"
      Top             =   1170
      Width           =   2220
   End
   Begin VB.ComboBox cmbTahunPelajaran 
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Text            =   "cmbTahunPelajaran"
      Top             =   1170
      Width           =   2265
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   13920
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DATA JADWAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         TabIndex        =   24
         Top             =   405
         Width           =   13830
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   0
      TabIndex        =   22
      Top             =   5085
      Width           =   6990
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   465
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_ubah 
         Caption         =   "Ubah"
         Height          =   465
         Left            =   2655
         TabIndex        =   10
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
         Height          =   465
         Left            =   3915
         TabIndex        =   11
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_baru 
         Caption         =   "Baru"
         Height          =   465
         Left            =   1395
         TabIndex        =   9
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "Keluar"
         Height          =   465
         Left            =   5670
         TabIndex        =   12
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Angsuran"
      Height          =   3480
      Left            =   0
      TabIndex        =   19
      Top             =   1530
      Width           =   6990
      Begin VB.ComboBox cmbHari 
         Height          =   315
         Left            =   1485
         TabIndex        =   39
         Text            =   "cmbHari"
         Top             =   765
         Width           =   2535
      End
      Begin VB.CommandButton cmdAmbilKelas 
         Caption         =   "..."
         Height          =   330
         Left            =   6600
         TabIndex        =   36
         Top             =   1170
         Width           =   285
      End
      Begin VB.TextBox txtKelas 
         Height          =   350
         Left            =   1485
         TabIndex        =   35
         Text            =   "txtKelas"
         Top             =   1170
         Width           =   5100
      End
      Begin VB.CommandButton cmdAmbilIdPelajaran 
         Caption         =   "..."
         Height          =   330
         Left            =   6600
         TabIndex        =   34
         Top             =   1665
         Width           =   285
      End
      Begin VB.TextBox txtPelajaran 
         Height          =   350
         Left            =   1485
         TabIndex        =   33
         Text            =   "txtPelajaran"
         Top             =   1665
         Width           =   5100
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "..."
         Height          =   330
         Left            =   6600
         TabIndex        =   5
         Top             =   2115
         Width           =   285
      End
      Begin VB.TextBox txtGuru 
         Height          =   350
         Left            =   1485
         TabIndex        =   4
         Text            =   "txtGuru"
         Top             =   2115
         Width           =   5100
      End
      Begin MSComCtl2.DTPicker DTPickerWaktuMulai 
         Height          =   375
         Left            =   1485
         TabIndex        =   6
         Top             =   2520
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   116916227
         CurrentDate     =   42566
      End
      Begin VB.TextBox txtID 
         Height          =   350
         Left            =   1500
         TabIndex        =   3
         Text            =   "txtID"
         Top             =   315
         Width           =   2500
      End
      Begin MSComCtl2.DTPicker DTPickerWaktuSelesai 
         Height          =   375
         Left            =   4635
         TabIndex        =   7
         Top             =   2520
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   116916227
         CurrentDate     =   42566
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   810
         Width           =   285
      End
      Begin VB.Label lblIDKelas 
         Caption         =   "llbIDKelas"
         Height          =   375
         Left            =   4140
         TabIndex        =   37
         Top             =   315
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblIdGuru 
         Caption         =   "lblIdGuru"
         Height          =   330
         Left            =   5625
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Guru"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label Label10 
         Caption         =   "Jam Selesai"
         Height          =   330
         Left            =   3330
         TabIndex        =   30
         Top             =   2565
         Width           =   1230
      End
      Begin VB.Label Label9 
         Caption         =   "Jam Mulai"
         Height          =   330
         Left            =   180
         TabIndex        =   29
         Top             =   2565
         Width           =   1230
      End
      Begin VB.Label lblIdPelajaran 
         Caption         =   "lblIDmatapelajaran"
         Height          =   330
         Left            =   4635
         TabIndex        =   28
         Top             =   315
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pelajaran"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1710
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         Height          =   345
         Left            =   180
         TabIndex        =   21
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
         Height          =   345
         Left            =   180
         TabIndex        =   20
         Top             =   1215
         Width           =   1290
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Semester"
      Height          =   4785
      Left            =   7065
      TabIndex        =   0
      Top             =   1215
      Width           =   6810
      Begin VB.CommandButton cmdCariDataJadwal 
         Caption         =   "Cari"
         Height          =   375
         Left            =   5265
         TabIndex        =   14
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox txtCariDataJadwal 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Text            =   "txtCariDataJadwal"
         Top             =   300
         Width           =   3150
      End
      Begin MSComctlLib.ListView lstJadwal 
         Height          =   3525
         Left            =   90
         TabIndex        =   15
         Top             =   810
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   6218
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Pencarian"
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Jumlah Pelajaran: "
         Height          =   255
         Left            =   135
         TabIndex        =   17
         Top             =   4365
         Width           =   1365
      End
      Begin VB.Label lblJumlahPelajaran 
         Caption         =   "lblJumlahGuru"
         Height          =   255
         Left            =   1530
         TabIndex        =   16
         Top             =   4365
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   465
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   820
      _Version        =   393216
      Format          =   116916225
      CurrentDate     =   42569
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Semester"
      Height          =   195
      Left            =   3555
      TabIndex        =   26
      Top             =   1215
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tahun Pelajaran"
      Height          =   195
      Left            =   0
      TabIndex        =   25
      Top             =   1215
      Width           =   1170
   End
End
Attribute VB_Name = "frmListJadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public readyForm As Boolean

Sub FormEdit()
    EnableTextBox Me
    Clear Me

    cmd_simpan.Enabled = False
    cmd_baru.Enabled = True
    cmd_ubah.Enabled = True
    cmd_hapus.Enabled = True
    txtID.Enabled = False
    cmd_baru.Caption = "Batal"
End Sub

Sub FormBaru()
    Dim tahun As String
    Dim semester As String

    tahun = cmbTahunPelajaran.Text
    semester = cmbSemester.Text

    EnableTextBox Me
    Clear Me

    cmd_simpan.Enabled = True
    cmd_baru.Enabled = True
    cmd_hapus.Enabled = False
    cmd_ubah.Enabled = False

    cmbTahunPelajaran.Text = tahun
    cmbSemester.Text = semester

    cmd_baru.Caption = "Batal"
    txtID.Enabled = False

End Sub

Private Sub cmbSemester_Change()
    If readyForm Then
        ListJadwal Me.txtCariDataJadwal.Text
    End If
End Sub

Private Sub cmbTahunPelajaran_Change()
    If readyForm Then
        ListJadwal Me.txtCariDataJadwal.Text
    End If
End Sub

Private Sub cmd_baru_Click()
    If cmd_baru.Caption = "Baru" Then
        FormBaru
    Else
        DisableBaru
    End If
End Sub

Sub SystemStarup()
    Dim rs As ADODB.Recordset
    cmbSemester.Clear
    cmbTahunPelajaran.Clear
    cmbHari.Clear

    cmbHari.AddItem "Senin"
    cmbHari.AddItem "Selasa"
    cmbHari.AddItem "Rabu"
    cmbHari.AddItem "Kamis"
    cmbHari.AddItem "Jumat"
    cmbHari.AddItem "Sabtu"
    cmbHari.AddItem "Minggu"


    Set rs = openRecordset("select * from semester")

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cmbSemester.AddItem rs.Fields("nama")
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)

    Set rs = openRecordset("select distinct(tahun) from jadwal")

    cmbTahunPelajaran.Text = CStr(DTPicker1.year)

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cmbTahunPelajaran.AddItem rs.Fields("tahun")
            rs.MoveNext
        Loop
    End If

    Dim dataFile() As String
    Dim i As Integer

    dataFile = loadDataFromFile("\conf.dat")

    cmbSemester.Text = getDataFromArray(dataFile, "semester")
    cmbTahunPelajaran.Text = getDataFromArray(dataFile, "tahun")

    Call closeRecordset(rs)

    ListJadwal txtCariDataJadwal.Text
End Sub

Sub DisableBaru()
    DisableTextBox Me
    Clear Me

    cmd_simpan.Enabled = False
    cmd_baru.Enabled = True
    cmd_hapus.Enabled = False
    cmd_ubah.Enabled = False

    cmbTahunPelajaran.Enabled = True
    cmbSemester.Enabled = True

    txtCariDataJadwal.Enabled = True

    cmd_baru.Caption = "Baru"
    txtID.Enabled = False
    ListJadwal ""
End Sub

Public Sub loadData(parameter As String)
    Dim tahun As String
    Dim semester As String

    tahun = cmbTahunPelajaran.Text
    semester = cmbSemester.Text

    FormEdit
    txtID.Text = parameter
    cmbTahunPelajaran.Text = tahun
    cmbSemester.Text = semester

    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset

    Set rs = openRecordset("select * from jadwal where id = '" + Trim$(parameter) + "'")

    If Not rs.EOF Then
        rs.MoveFirst
        Me.txtID.Text = rs.Fields("id")
        cmbHari.Text = HariDariUrutanTanggal(CInt(rs.Fields("hari")))

        Me.lblIdGuru.Caption = rs.Fields("idguru")

        Set rs1 = openRecordset("select * from guru where id = '" & Trim$(rs.Fields("idguru")) & "'")

        If Not rs1.EOF Then
            txtGuru.Text = rs1.Fields("id") & " - " & rs1.Fields("nama")
        End If

        Me.lblIDKelas.Caption = rs.Fields("idkelas")

        Call closeRecordset(rs1)

        Set rs1 = openRecordset("select * from kelas where id = '" & Trim$(rs.Fields("idkelas")) & "'")

        If Not rs1.EOF Then
            txtKelas.Text = rs1.Fields("id") & " - " & rs1.Fields("nama")
        End If

        Me.lblIdPelajaran.Caption = rs.Fields("idpelajaran")

        Call closeRecordset(rs1)

        Set rs1 = openRecordset("select * from pelajaran where id = '" & Trim$(rs.Fields("idpelajaran")) & "'")

        If Not rs1.EOF Then
            txtPelajaran.Text = rs1.Fields("id") & " - " & rs1.Fields("nama")
        End If

        Call closeRecordset(rs1)
    End If


    Call closeRecordset(rs)
End Sub

Sub Hapus(parameter As String)
    If parameter = "" Then
        Clear Me
        Exit Sub
    End If

    If MsgBox("Apakah anda yakin untuk menghapus data ini?", vbOKCancel, "Konfirmasi") = vbOK Then
        Dim rs As ADODB.Recordset

        Set rs = openRecordset("select * from jadwal where id = '" + Trim$(parameter) + "'")

        If Not rs.EOF Then
            rs!deleted = "1"
            rs.Update
        End If

        Call closeRecordset(rs)
    End If

    DisableBaru
    SystemStarup
End Sub

Sub Simpan(parameter As String, no As String)
    'If txtID.Text = "" Then GoTo ErrorNull
    'If Len(txt_noid.Text) > 11 Then GoTo ErrorLengthNoid
    'If Len(txt_telepon.Text) > 13 Then GoTo ErrorLengthTelepon
    'If Len(txt_nama.Text) > 20 Then GoTo ErrorLengthNama

    Dim rs As ADODB.Recordset

    'RS1.CursorLocation = adUseClient
    'RS1.Open "select tahunpelajaran.id from tahunpelajaran inner join semester on semester.id = tahunpelajaran.semester where tahun = '" & Trim$(cmbTahunPelajaran.Text) & "' and nama = '" & Trim$(cmbSemester.Text) & "'")
    '
    'Dim id As String
    '
    'If Not RS1.EOF Then
    '    id = RS1.Fields("id")
    'Else

    Set rs = openRecordset("select * from semester where nama = '" & Trim$(cmbSemester.Text) & "'")

    Dim semester As String

    If Not rs.EOF Then
        semester = rs.Fields("id")
    End If

    Call closeRecordset(rs)

    '
    '    RS2.CursorLocation = adUseClient
    '    RS2.Open "select * from tahunpelajaran")
    '
    '    RS2.AddNew
    '    RS2!tahun = Trim$(cmbTahunPelajaran.Text)
    '    RS2!semester = semester
    '    RS2.Update
    '
    '    RS3.CursorLocation = adUseClient
    '    RS3.Open "select tahunpelajaran.id from tahunpelajaran inner join semester on semester.id = tahunpelajaran.semester where tahun = '" & Trim$(cmbTahunPelajaran.Text) & "' and nama = '" & Trim$(cmbSemester.Text) & "'")
    '
    '    If Not RS3.EOF Then
    '        id = RS3.Fields("id")
    '    End If
    'End If


    Dim Param As String

    If parameter = "edit" Then
        Param = "select * from jadwal where id = '" + Trim$(no) + "'"
    Else
        Param = "select * from jadwal"
    End If


    Set rs = openRecordset(Param)

    If parameter = "new" Then
        rs.AddNew
    End If

    Dim hari As Integer

    Select Case cmbHari.Text
        Case "Senin"
            hari = 2
        Case "Selasa"
            hari = 3
        Case "Rabu"
            hari = 4
        Case "Kamis"
            hari = 5
        Case "Jumat"
            hari = 6
        Case "Sabtu"
            hari = 7
    End Select

    rs!semester = semester
    rs!idkelas = Trim$(lblIDKelas.Caption)
    rs!hari = hari
    rs!tahun = cmbTahunPelajaran.Text
    rs!idguru = Trim$(lblIdGuru.Caption)
    rs!idpelajaran = Trim$(lblIdPelajaran.Caption)
    rs!waktumulai = CStr(Format(DTPickerWaktuMulai.value, "yyyy-mm-dd HH:mm:ss"))
    rs!waktuselesai = CStr(Format(DTPickerWaktuSelesai.value, "yyyy-mm-dd HH:mm:ss"))
    rs.Update

    Call closeRecordset(rs)

    If parameter = "new" Then
        MsgBox "Data Berhasil Tersimpan"
    Else
        MsgBox "Data Berhasil Terubah"
    End If

    DisableBaru
    SystemStarup

    Exit Sub
ErrorNull:
    MsgBox "Field masih ada yang kosong !"
    Exit Sub
End Sub

Function jumlahPelajaran() As Integer
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select count(*) as jumlah from jadwal")

    If Not rs.EOF Then
        rs.MoveFirst
        JumlahGuru = rs.Fields("jumlah")
    Else
        JumlahGuru = 0
    End If

    Call closeRecordset(rs)
End Function

Sub ListJadwal(parameter As String)
    With Me.lstJadwal

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Hari", 500
        .ColumnHeaders.Add , , "Jadwal Mulai", 2000
        .ColumnHeaders.Add , , "Jadwal Selesai", 500
        .ColumnHeaders.Add , , "Id Pelajaran", 500
        .ColumnHeaders.Add , , "Id Kelas", 2000
        .ColumnHeaders.Add , , "Id Guru", 500

        Dim rs As ADODB.Recordset

        'RS.Open "select jadwal.*, convert(varchar(20),waktumulai,8) as mulai, convert(varchar(20),waktuselesai,8) as selesai from jadwal inner join tahunpelajaran on tahunpelajaran.id = jadwal.semester inner join semester on tahunpelajaran.semester = semester.id where jadwal.id like '%" + Trim$(parameter) + "%' and tahun = '" & Trim$(cmbTahunPelajaran.Text) & "' and semester.nama = '" + Trim$(cmbSemester.Text) + "'")
        Set rs = openRecordset("select jadwal.*, convert(varchar(20),waktumulai,8) as mulai, convert(varchar(20),waktuselesai,8) as selesai from jadwal inner join guru on guru.id = jadwal.idguru inner join pelajaran on pelajaran.id = jadwal.idpelajaran inner join semester on jadwal.semester = semester.id where jadwal.id like '%" + Trim$(parameter) + "%' and tahun = '" & Trim$(cmbTahunPelajaran.Text) & "' and semester.nama = '" + Trim$(cmbSemester.Text) + "'  and jadwal.deleted = 0 and guru.deleted = 0 and pelajaran.deleted = 0")

        .ListItems.Clear

        Dim i As Integer
        i = 1

        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Dim List As ListItem
                Set List = .ListItems.Add(, , i)

                fillListView List, rs, "id"
                fillListView List, rs, "hari"
                fillListView List, rs, "mulai"
                fillListView List, rs, "selesai"
                fillListView List, rs, "idpelajaran"
                fillListView List, rs, "idkelas"
                fillListView List, rs, "idguru"

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)

    End With

    lblJumlahPelajaran.Caption = jumlahPelajaran
End Sub

Private Sub cmd_hapus_Click()
    'Hapus txtID.Text
    MsgBox CStr(Format(DTPickerWaktuMulai.value, "yyyy-mm-dd HH:mm:ss"))

End Sub

Private Sub cmd_keluar_Click()
    Unload Me
End Sub

Private Sub cmd_simpan_Click()
    Simpan "new", "0"
End Sub

Private Sub cmd_ubah_Click()
    Simpan "edit", txtID.Text
End Sub

Private Sub cmdAmbilIdPelajaran_Click()
    frmDataPelajaran.Show , Me
End Sub

Private Sub cmdAmbilKelas_Click()
    frmDataKelas.Show , Me
End Sub

Private Sub Command1_Click()
    frmDataGuru.Show , Me
End Sub

Private Sub Form_Load()
    readyForm = False
    Clear Me
    DisableBaru
    Me.SystemStarup
    Me.readyForm = False
End Sub

Private Sub lstJadwal_Click()
    With Me.lstJadwal
        If .ListItems.Count = 0 Then Exit Sub
        loadData .SelectedItem.SubItems(1)
        DTPickerWaktuMulai.value = CStr(Format$(.SelectedItem.SubItems(3), "hh:mm:ss"))
        DTPickerWaktuSelesai.value = CStr(Format$(.SelectedItem.SubItems(4), "hh:mm:ss"))
    End With
End Sub


