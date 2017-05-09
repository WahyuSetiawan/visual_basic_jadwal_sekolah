VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPengaturan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   45
      TabIndex        =   5
      Top             =   3105
      Width           =   8745
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   195
         Left            =   5895
         TabIndex        =   9
         Top             =   450
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   344
         _Version        =   393216
         Format          =   100925441
         CurrentDate     =   42574
      End
      Begin VB.CommandButton cmd_simpan 
         Caption         =   "Simpan"
         Height          =   465
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_Batal 
         Caption         =   "Batal"
         Height          =   465
         Left            =   1395
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmd_keluar 
         Caption         =   "Keluar"
         Height          =   465
         Left            =   7425
         TabIndex        =   6
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Pengaturan"
      Height          =   3030
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8730
      Begin VB.TextBox txtHost 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1305
         Width           =   3030
      End
      Begin VB.TextBox txtUsername 
         Height          =   330
         Left            =   1440
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1710
         Width           =   5010
      End
      Begin VB.TextBox txtPassword 
         Height          =   330
         Left            =   1440
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2115
         Width           =   5010
      End
      Begin VB.TextBox txtDatabse 
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2520
         Width           =   3030
      End
      Begin VB.TextBox txtPORT 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Text            =   "txtport"
         Top             =   900
         Width           =   3030
      End
      Begin VB.ComboBox cmbTahunAktif 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   495
         Width           =   1770
      End
      Begin VB.ComboBox cmbSemesterAktif 
         Height          =   315
         Left            =   4455
         TabIndex        =   1
         Text            =   "Combo2"
         Top             =   495
         Width           =   3930
      End
      Begin VB.Label Label5 
         Caption         =   "Host"
         Height          =   330
         Left            =   225
         TabIndex        =   19
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
         Height          =   330
         Left            =   225
         TabIndex        =   18
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   330
         Left            =   225
         TabIndex        =   17
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Database"
         Height          =   330
         Left            =   225
         TabIndex        =   16
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Port"
         Height          =   330
         Left            =   225
         TabIndex        =   10
         Top             =   900
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tahun Aktif"
         Height          =   225
         Left            =   225
         TabIndex        =   4
         Top             =   495
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Semester Aktif"
         Height          =   195
         Left            =   3330
         TabIndex        =   3
         Top             =   540
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Simpan(parameter As String, value As String)

End Sub

Sub systemLoad()
    Clear Me

    cmbSemesterAktif.Clear
    cmbTahunAktif.Clear

    Dim rs As ADODB.Recordset
    Set rs = openRecordset("select * from semester")

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cmbSemesterAktif.AddItem rs.Fields("nama")
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)


    Dim dataFile() As String
    Dim i As Integer

    dataFile = loadDataFromFile("\conf.dat")

    cmbSemesterAktif.Text = getDataFromArray(dataFile, "semester")
    cmbTahunAktif.Text = getDataFromArray(dataFile, "tahun")
    txtPORT.Text = getDataFromArray(dataFile, "port")

    Dim file As String

    file = "\connectionconf.dat"

    dataFile = loadDataFromFile(file)

    txtHost.Text = getDataFromArray(dataFile, "host")
    txtUsername.Text = getDataFromArray(dataFile, "username")
    txtPassword.Text = getDataFromArray(dataFile, "password")
    txtDatabse.Text = getDataFromArray(dataFile, "database")
End Sub

Private Sub cmd_Batal_Click()
    Clear Me
    systemLoad
End Sub

Private Sub cmd_keluar_Click()
    Visible = False
End Sub


Private Sub cmd_simpan_Click()
    Dim data() As String

    ReDim Preserve data(0)
    data(0) = "tahun:" + cmbTahunAktif.Text
    ReDim Preserve data(1)
    data(1) = "semester:" + cmbSemesterAktif.Text
    ReDim Preserve data(2)
    data(2) = "port:" + txtPORT.Text

    createAndSaveFile pathFileConfiguration, data

    If testKonekToServer(txtHost.Text, txtUsername.Text, txtPassword.Text, txtDatabse.Text) Then
        MsgBox "koneksi sukses"

        Dim dataFile() As String
        Dim i As Integer

        Dim file As String

        file = "\connectionconf.dat"

        Dim host As String
        Dim username As String
        Dim password As String
        Dim database As String

        ReDim Preserve data(0)
        data(0) = "host:" & txtHost.Text
        ReDim Preserve data(1)
        data(1) = "username:" & txtUsername.Text
        ReDim Preserve data(2)
        data(2) = "password:" & txtPassword.Text
        ReDim Preserve data(3)
        data(3) = "database:" & txtDatabse.Text

        createAndSaveFile file, data
    Else
        MsgBox "koneksi gagal"
    End If

End Sub

Private Sub Form_Load()
    Me.systemLoad
End Sub


