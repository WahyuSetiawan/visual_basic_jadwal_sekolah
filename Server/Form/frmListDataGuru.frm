VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListDataGuru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guru"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   19035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Data Guru"
      Height          =   6540
      Left            =   90
      TabIndex        =   25
      Top             =   1215
      Width           =   8730
      Begin VB.ComboBox cmbJabatan 
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   2205
         Width           =   3390
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   1485
         TabIndex        =   5
         Text            =   "Combo2"
         Top             =   2205
         Width           =   2580
      End
      Begin VB.TextBox txtID 
         Height          =   350
         Left            =   1485
         TabIndex        =   1
         Text            =   "txtID"
         Top             =   315
         Width           =   2500
      End
      Begin VB.ComboBox cmbAgama 
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   2655
         Width           =   2580
      End
      Begin VB.ComboBox cmbJenisKelamin 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1215
         Width           =   3120
      End
      Begin VB.TextBox txtTempat 
         Height          =   375
         Left            =   1500
         TabIndex        =   8
         Text            =   "txtTempat"
         Top             =   3150
         Width           =   3840
      End
      Begin VB.TextBox txtNip 
         Height          =   375
         Left            =   1500
         TabIndex        =   4
         Text            =   "txtNip"
         Top             =   1710
         Width           =   4740
      End
      Begin VB.TextBox txtNama 
         Height          =   375
         Left            =   1500
         TabIndex        =   2
         Text            =   "txtNama"
         Top             =   720
         Width           =   6990
      End
      Begin MSComCtl2.DTPicker dtpTanggalLahir 
         Height          =   330
         Left            =   1500
         TabIndex        =   9
         Top             =   3645
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   582
         _Version        =   393216
         Format          =   101122049
         CurrentDate     =   42562
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
         Height          =   195
         Left            =   4230
         TabIndex        =   34
         Top             =   2250
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   315
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   3690
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tempat"
         Height          =   195
         Left            =   195
         TabIndex        =   31
         Top             =   3195
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agama"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   2700
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   2205
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   1755
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   195
         Left            =   195
         TabIndex        =   27
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama "
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   765
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   45
      TabIndex        =   20
      Top             =   45
      Width           =   18960
      Begin VB.Label lbl_nopinjaman 
         Caption         =   "Label7"
         Height          =   600
         Left            =   585
         TabIndex        =   22
         Top             =   315
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "DATA GURU"
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
         TabIndex        =   21
         Top             =   360
         Width           =   18825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   90
      TabIndex        =   19
      Top             =   7830
      Width           =   6990
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   465
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_ubah 
         Caption         =   "Ubah"
         Height          =   465
         Left            =   2655
         TabIndex        =   12
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
         Height          =   465
         Left            =   3915
         TabIndex        =   13
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_baru 
         Caption         =   "Baru"
         Height          =   465
         Left            =   1395
         TabIndex        =   11
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "Keluar"
         Height          =   465
         Left            =   5670
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian Data Guru"
      Height          =   7485
      Left            =   8910
      TabIndex        =   0
      Top             =   1215
      Width           =   10095
      Begin VB.TextBox txtCariDataGuru 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Text            =   "txtCariDataGuru"
         Top             =   300
         Width           =   6255
      End
      Begin VB.CommandButton cmdDataGuru 
         Caption         =   "Cari"
         Height          =   375
         Left            =   8580
         TabIndex        =   16
         Top             =   300
         Width           =   1395
      End
      Begin MSComctlLib.ListView lstGuru 
         Height          =   5955
         Left            =   90
         TabIndex        =   18
         Top             =   810
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   10504
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblJumlahGuru 
         Caption         =   "lblJumlahGuru"
         Height          =   255
         Left            =   1215
         TabIndex        =   24
         Top             =   6930
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Jumlah Guru : "
         Height          =   255
         Left            =   135
         TabIndex        =   23
         Top             =   6930
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Pencarian"
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmListDataGuru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub FormEdit()
    EnableTextBox Me
    Clear Me

    cmd_simpan.Enabled = False
    cmd_baru.Enabled = True
    cmd_ubah.Enabled = True
    cmd_hapus.Enabled = True
    txtID.Enabled = False
    cmd_baru.Caption = "Batal"

    Me.txtCariDataGuru.Enabled = True
End Sub

Sub FormBaru()
    EnableTextBox Me
    Clear Me

    cmd_simpan.Enabled = True
    cmd_baru.Enabled = True
    cmd_hapus.Enabled = False
    cmd_ubah.Enabled = False

    cmd_baru.Caption = "Batal"
    txtID.Enabled = False

    Me.txtCariDataGuru.Enabled = True

    ListGuru ""
End Sub

Sub DisableBaru()
    DisableTextBox Me
    Clear Me

    cmd_simpan.Enabled = False
    cmd_baru.Enabled = True
    cmd_hapus.Enabled = False
    cmd_ubah.Enabled = False

    cmd_baru.Caption = "Baru"
    txtID.Enabled = False

    Me.txtCariDataGuru.Enabled = True
    ListGuru ""
End Sub

Public Sub LoadGuru(parameter As String)
    FormEdit
    txtID.Text = parameter

    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select * from guru where id = '" + Trim$(parameter) + "'")

    If Not rs.EOF Then
        rs.MoveFirst
        Me.txtNama.Text = rs.Fields("Nama")
        Me.cmbJenisKelamin.Text = rs.Fields("jeniskelamin")
        Me.txtNip.Text = rs.Fields("nip")
        Me.cmbAgama.Text = rs.Fields("agama")
        Me.txtTempat.Text = rs.Fields("tempat")
        Me.dtpTanggalLahir.value = DateValue(rs.Fields("tanggallahir"))
        Me.cmbJabatan.Text = rs.Fields("jabatan")
        Me.cmbStatus.Text = DateValue(rs.Fields("tanggallahir"))

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

        Set rs = openRecordset("select * from guru where id = '" + Trim$(parameter) + "'")

        If Not rs.EOF Then
            rs!deleted = "1"
            rs.Update
        End If
        Call closeRecordset(rs)
    End If

    DisableBaru

End Sub

Sub Simpan(parameter As String, no As String)
    If txtNama.Text = "" Then GoTo ErrorNull
    If cmbJenisKelamin.Text = "" Then GoTo ErrorNull
    'If txtNip.Text = "" Then GoTo ErrorNull
    If cmbStatus.Text = "" Then GoTo ErrorNull
    If cmbAgama.Text = "" Then GoTo ErrorNull
    If txtTempat.Text = "" Then GoTo ErrorNull
    If dtpTanggalLahir.value = "" Then GoTo ErrorNull
    If cmbJabatan.Text = "" Then GoTo ErrorNull
    'If Len(txt_noid.Text) > 11 Then GoTo ErrorLengthNoid
    'If Len(txt_telepon.Text) > 13 Then GoTo ErrorLengthTelepon
    'If Len(txt_nama.Text) > 20 Then GoTo ErrorLengthNama

    Dim Param As String
    Dim rs As ADODB.Recordset

    If parameter = "edit" Then
        Param = "select * from guru where id = '" + Trim$(no) + "'"
    Else
        Param = "select * from guru"
    End If

    Set rs = openRecordset(Param)

    If parameter = "new" Then
        rs.AddNew
    End If

    rs!nama = Trim$(Me.txtNama.Text)
    rs!jeniskelamin = Trim$(Me.cmbJenisKelamin.Text)
    rs!nip = Trim$(Me.txtNip.Text)
    rs!Status = Trim$(Me.cmbStatus.Text)
    rs!agama = Trim$(Me.cmbAgama.Text)
    rs!tempat = Trim$(Me.txtTempat.Text)
    rs!tanggallahir = Trim$(Me.dtpTanggalLahir.value)
    rs!jabatan = Trim$(Me.cmbJabatan.Text)
    rs.Update

    Call closeRecordset(rs)

    MsgBox "Data Berhasil Tersimpan"

    DisableBaru
    Exit Sub
ErrorNull:
    MsgBox "Field masih ada yang kosong !"
    Exit Sub
ErrorLengthNoid:
    MsgBox "Field no anggota max 11 !"
    txt_noid.SetFocus
    Exit Sub
ErrorLengthNama:
    MsgBox "Field nama anggota max 20 !"
    txt_nama.SetFocus
    Exit Sub
ErrorLengthTelepon:
    MsgBox "Field telepon max 4 !"
    txt_telepon.SetFocus
    Exit Sub

MaxLengthPassword:
    MsgBox "Maksimum password hanya 15 karakter"
    txtPassword.SetFocus
    Exit Sub
End Sub


Sub SystemStartup()
    cmbAgama.Clear
    cmbAgama.Text = "Islam"
    cmbAgama.AddItem "Islam"
    cmbAgama.AddItem "Kristen"
    cmbAgama.AddItem "Prostestan"
    cmbAgama.AddItem "Hindu"
    cmbAgama.AddItem "Buddha"

    cmbJenisKelamin.Clear
    cmbJenisKelamin.Text = "L"
    cmbJenisKelamin.AddItem "L"
    cmbJenisKelamin.AddItem "P"

    cmbJabatan.Clear
    cmbStatus.Clear

    Dim rs As ADODB.Recordset


    Set rs = openRecordset("select distinct(jabatan) as jabatan from guru")

    If Not rs.EOF Then
        rs.MoveFirst
        cmbJabatan.Text = rs.Fields("jabatan")
        Do While Not rs.EOF
            cmbJabatan.AddItem rs.Fields("jabatan")
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)


    Set rs = openRecordset("select distinct(status) as status from guru")

    If Not rs.EOF Then
        rs.MoveFirst
        cmbStatus.Text = rs.Fields("status")
        Do While Not rs.EOF
            cmbStatus.AddItem rs.Fields("status")
            rs.MoveNext
        Loop
    End If
    Call closeRecordset(rs)
End Sub
Function JumlahGuru() As Integer
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select count(*) as jumlah from Guru")

    If Not rs.EOF Then
        rs.MoveFirst
        JumlahGuru = rs.Fields("jumlah")
    Else
        JumlahGuru = 0
    End If

    Call closeRecordset(rs)
End Function

Sub ListGuru(parameter As String)
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

        Set rs = openRecordset("select * from Guru where nama like '%" + Trim$(parameter) + "%'  and deleted = 0")

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

    lblJumlahGuru.Caption = JumlahGuru
End Sub

Private Sub cmd_baru_Click()
    If cmd_baru.Caption = "Baru" Then
        FormBaru
    Else
        DisableBaru
    End If
End Sub

Private Sub cmd_hapus_Click()
    Hapus txtID.Text
    Clear Me
End Sub

Private Sub cmd_keluar_Click()
    Visible = False
End Sub

Private Sub cmd_simpan_Click()
    Simpan "new", "0"
    Clear Me
End Sub

Private Sub cmd_ubah_Click()
    Simpan "edit", txtID.Text
    Clear Me
End Sub

Private Sub cmdDataGuru_Click()
    Me.ListGuru txtCariDataGuru.Text
End Sub

Private Sub cmdTambahGuru_Click()
    frmDataGuru.Show , Me
End Sub

Private Sub Form_Load()
    Clear Me
    Me.ListGuru txtCariDataGuru.Text
    Me.SystemStartup
    DisableBaru
End Sub

Private Sub lstGuru_Click()
    With Me.lstGuru
        If .ListItems.Count = 0 Then Exit Sub
        LoadGuru .SelectedItem.SubItems(1)
    End With
End Sub


