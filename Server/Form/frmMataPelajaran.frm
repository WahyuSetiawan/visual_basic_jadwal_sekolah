VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListMataPelajaran 
   Caption         =   "Mata Pelajaran"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Data Semester"
      Height          =   4245
      Left            =   7065
      TabIndex        =   13
      Top             =   1215
      Width           =   6810
      Begin VB.TextBox txtCariDataMataPelajaran 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Text            =   "txtCariDataKelas"
         Top             =   300
         Width           =   3150
      End
      Begin VB.CommandButton cmdCariDataMataPelajaran 
         Caption         =   "Cari"
         Height          =   375
         Left            =   5265
         TabIndex        =   14
         Top             =   270
         Width           =   1395
      End
      Begin MSComctlLib.ListView lstMataPelajaran 
         Height          =   2985
         Left            =   90
         TabIndex        =   16
         Top             =   810
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   5265
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblJumlahPelajaran 
         Caption         =   "lblJumlahGuru"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3825
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Jumlah Pelajaran: "
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   3825
         Width           =   1365
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
   Begin VB.Frame Frame1 
      Caption         =   "Data Angsuran"
      Height          =   3300
      Left            =   0
      TabIndex        =   8
      Top             =   1215
      Width           =   6990
      Begin VB.TextBox txtNama 
         Height          =   350
         Left            =   1500
         TabIndex        =   10
         Text            =   "txtNama"
         Top             =   855
         Width           =   5100
      End
      Begin VB.TextBox txtID 
         Height          =   350
         Left            =   1500
         TabIndex        =   9
         Text            =   "txtID"
         Top             =   315
         Width           =   2500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   855
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         Height          =   345
         Left            =   180
         TabIndex        =   11
         Top             =   315
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   4545
      Width           =   6990
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "Keluar"
         Height          =   465
         Left            =   5670
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_baru 
         Caption         =   "Baru"
         Height          =   465
         Left            =   1395
         TabIndex        =   6
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus"
         Height          =   465
         Left            =   3915
         TabIndex        =   5
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_ubah 
         Caption         =   "Ubah"
         Height          =   465
         Left            =   2655
         TabIndex        =   4
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   465
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13920
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DATA MATA PELAJARAN"
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
         TabIndex        =   1
         Top             =   405
         Width           =   13830
      End
   End
End
Attribute VB_Name = "frmListMataPelajaran"
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

    txtCariDataMataPelajaran.Enabled = True
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
    Me.ListPelajaran txtCariDataMataPelajaran.Text

    txtCariDataMataPelajaran.Enabled = True
End Sub

Private Sub cmd_baru_Click()
    If cmd_baru.Caption = "Baru" Then
        FormBaru
    Else
        DisableBaru
    End If
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
    ListPelajaran ""

    txtCariDataMataPelajaran.Enabled = True

    Me.txtCariDataMataPelajaran.Enabled = True
End Sub

Public Sub loadData(parameter As String)
    FormEdit
    txtID.Text = parameter

    Dim rs As ADODB.Recordset


    Set rs = openRecordset("select * from pelajaran where id = '" + Trim$(parameter) + "'")

    If Not rs.EOF Then
        rs.MoveFirst
        Me.txtNama.Text = rs.Fields("Nama")
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


        Set rs = openRecordset("select * from pelajaran where id = '" + Trim$(parameter) + "'")

        If Not rs.EOF Then
            rs!deleted = "1"
            rs.Update
        End If

        Call closeRecordset(rs)
    End If

    DisableBaru
    loadData txtCariDataMataPelajaran.Text
End Sub

Sub Simpan(parameter As String, no As String)
    If txtNama.Text = "" Then GoTo ErrorNull
    'If Len(txt_noid.Text) > 11 Then GoTo ErrorLengthNoid
    'If Len(txt_telepon.Text) > 13 Then GoTo ErrorLengthTelepon
    'If Len(txt_nama.Text) > 20 Then GoTo ErrorLengthNama

    Dim rs As ADODB.Recordset

    Dim Param As String

    If parameter = "edit" Then
        Param = "select * from pelajaran where id = '" + Trim$(no) + "'"
    Else
        Param = "select * from pelajaran"
    End If


    Set rs = openRecordset(Param)

    If parameter = "new" Then
        rs.AddNew
    End If

    rs!nama = Trim$(Me.txtNama.Text)
    rs.Update

    Call closeRecordset(rs)

    If parameter = "new" Then
        MsgBox "Data Berhasil Tersimpan"
    Else
        MsgBox "Data Berhasil Terubah"
    End If

    DisableBaru
    Me.ListPelajaran txtCariDataMataPelajaran.Text

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
End Sub

Function jumlahPelajaran() As Integer
    Dim rs As ADODB.Recordset


    Set rs = openRecordset("select count(*) as jumlah from pelajaran")


    If Not rs.EOF Then
        rs.MoveFirst
        JumlahGuru = rs.Fields("jumlah")
    Else
        JumlahGuru = 0
    End If

    Call closeRecordset(rs)
End Function

Sub ListPelajaran(parameter As String)
    With Me.lstMataPelajaran

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Nama", 2000

        Dim rs As ADODB.Recordset


        Set rs = openRecordset("select * from pelajaran where nama like '%" + Trim$(parameter) + "%'  and deleted = 0")

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

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)
    End With

    lblJumlahPelajaran.Caption = jumlahPelajaran
End Sub

Private Sub cmdTambahMataPelajaran_Click()
    frmDataPelajaran.Show , Me
End Sub

Private Sub Command2_Click()
    Me.ListPelajaran Text1.Text
End Sub

Private Sub cmd_hapus_Click()
    Hapus txtID.Text
End Sub

Private Sub cmd_keluar_Click()
    Visible = False
End Sub

Private Sub cmd_simpan_Click()
    Simpan "new", "0"
End Sub

Private Sub cmd_ubah_Click()
    Simpan "edit", txtID.Text
End Sub

Private Sub cmdCariDataMataPelajaran_Click()
    ListPelajaran txtCariDataMataPelajaran.Text
End Sub

Private Sub Form_Load()
    Clear Me

    Me.ListPelajaran txtCariDataMataPelajaran.Text
    DisableBaru
End Sub

Private Sub lbl_nopinjaman_Click()

End Sub

Private Sub lstMataPelajaran_Click()
    With Me.lstMataPelajaran

        If .ListItems.Count = 0 Then Exit Sub
        loadData .SelectedItem.SubItems(1)
    End With
End Sub



