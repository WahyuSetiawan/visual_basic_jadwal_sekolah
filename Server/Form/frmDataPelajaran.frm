VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataPelajaran 
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Mencari Data Pelajaran"
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
Attribute VB_Name = "frmDataPelajaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

    Load frmListJadwal
    frmListJadwal.lblIdPelajaran = ListView1.SelectedItem.SubItems(1)
    frmListJadwal.txtPelajaran.Text = ListView1.SelectedItem.SubItems(1) + " - " + ListView1.SelectedItem.SubItems(2)
    frmListJadwal.Show
    Unload Me
End Sub





