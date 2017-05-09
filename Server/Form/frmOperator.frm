VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOperator 
   Caption         =   "Operator Management"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   90
      TabIndex        =   16
      Top             =   90
      Width           =   13920
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DATA OPERATOR"
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
         TabIndex        =   17
         Top             =   315
         Width           =   13830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   90
      TabIndex        =   10
      Top             =   4635
      Width           =   6990
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   465
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_ubah 
         Caption         =   "Ubah"
         Height          =   465
         Left            =   2655
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "Keluar"
         Height          =   465
         Left            =   5670
         TabIndex        =   11
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Operator"
      Height          =   3300
      Left            =   90
      TabIndex        =   5
      Top             =   1305
      Width           =   6990
      Begin VB.TextBox txtID 
         Height          =   350
         Left            =   1500
         TabIndex        =   7
         Text            =   "txtID"
         Top             =   315
         Width           =   2500
      End
      Begin VB.TextBox txtNama 
         Height          =   350
         Left            =   1500
         TabIndex        =   6
         Text            =   "txtNama"
         Top             =   855
         Width           =   5100
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   345
         Left            =   180
         TabIndex        =   9
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   855
         Width           =   1290
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Operator"
      Height          =   4245
      Left            =   7155
      TabIndex        =   0
      Top             =   1305
      Width           =   6810
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   5265
         TabIndex        =   2
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox txtCariDataOperator 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Text            =   "txtCariDataOPerator"
         Top             =   300
         Width           =   3150
      End
      Begin MSComctlLib.ListView lstOperator 
         Height          =   3255
         Left            =   90
         TabIndex        =   3
         Top             =   810
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOperator"
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
    cmd_hapus.Enabled = False
    txtID.Enabled = False
    cmd_baru.Caption = "Batal"
    txtCariDataOperator.Enabled = True
End Sub

Sub FormBaru()
    EnableTextBox Me
    Clear Me

    cmd_simpan.Enabled = True
    cmd_baru.Enabled = True
    cmd_hapus.Enabled = False
    cmd_ubah.Enabled = False

    cmd_baru.Caption = "Batal"
    txtID.Enabled = True
    ListSemester ""
    txtCariDataOperator.Enabled = True
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
    ListSemester ""
    txtCariDataOperator.Enabled = True
End Sub

Public Sub LoadDataSemester(parameter As String)
    FormEdit
    txtID.Text = parameter

    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select * from operator  where username = '" + Trim$(parameter) + "'")

    If Not rs.EOF Then
        rs.MoveFirst
        Me.txtID.Text = rs.Fields("username")
        Me.txtNama.Text = rs.Fields("pass")
    End If

    Call closeRecordset(rs)
End Sub

Sub HapusSemester(parameter As String)
    If parameter = "" Then
        Clear Me
        Exit Sub
    End If

    If MsgBox("Apakah anda yakin untuk menghapus data ini?", vbOKCancel, "Konfirmasi") = vbOK Then
        Dim rs As ADODB.Recordset


        Set rs = openRecordset("select * from operator where username = '" + Trim$(parameter) + "'")

        If Not rs.EOF Then
            rs.Delete adAffectCurrent
        End If

        Call closeRecordset(rs)
    End If
End Sub

Sub SimpanSemester(parameter As String, no As String)
    If txtNama.Text = "" Then GoTo ErrorNull
    'If Len(txt_noid.Text) > 11 Then GoTo ErrorLengthNoid
    'If Len(txt_telepon.Text) > 13 Then GoTo ErrorLengthTelepon
    'If Len(txt_nama.Text) > 20 Then GoTo ErrorLengthNama

    Dim rs As ADODB.Recordset

    Dim Param As String

    If parameter = "edit" Then
        Param = "select * from operator where username = '" + Trim$(no) + "'"
    Else
        Param = "select * from operator"
    End If


    Set rs = openRecordset(Param)

    If parameter = "new" Then
        rs.AddNew
    End If

    rs!username = Trim$(Me.txtID.Text)
    rs!pass = Trim$(Me.txtNama.Text)
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
End Sub

Sub ListSemester(parameter As String)
    With Me.lstOperator

        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .AllowColumnReorder = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 500
        .ColumnHeaders.Add , , "Username", 3000
        .ColumnHeaders.Add , , "Password", 2000

        Dim rs As ADODB.Recordset

        Set rs = openRecordset("select * from operator where username like '%" + Trim$(parameter) + "%'")

        .ListItems.Clear

        Dim i As Integer
        i = 1

        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Dim List As ListItem
                Set List = .ListItems.Add(, , i)

                fillListView List, rs, "username"
                fillListView List, rs, "pass"

                i = i + 1
                rs.MoveNext
            Loop
        End If

        Call closeRecordset(rs)
    End With
End Sub

Private Sub cmdTambahMataPelajaran_Click()
    frmDataKelas.Show , Me
End Sub


Private Sub cmd_baru_Click()
    If cmd_baru.Caption = "Baru" Then
        FormBaru
    Else
        DisableBaru
    End If
End Sub

Private Sub cmd_hapus_Click()
    HapusSemester txtID.Text
    ListSemester ""
End Sub

Private Sub cmd_keluar_Click()
    Visible = False
End Sub

Private Sub cmd_simpan_Click()
    SimpanSemester "new", "0"

    ListSemester ""
End Sub

Private Sub cmd_ubah_Click()
    SimpanSemester "edit", txtID.Text
    ListSemester ""
End Sub

Private Sub cmdDataGuru_Click()

End Sub

Private Sub cmdCari_Click()
    Me.ListSemester txtCariDataOperator.Text
End Sub

Private Sub Form_Load()
    Clear Me
    Me.ListSemester txtCariDataOperator.Text
    DisableBaru
End Sub


Private Sub lstOperator_DblClick()
    With Me.lstOperator

        If .ListItems.Count = 0 Then Exit Sub
        LoadDataSemester .SelectedItem.SubItems(1)

    End With
End Sub

