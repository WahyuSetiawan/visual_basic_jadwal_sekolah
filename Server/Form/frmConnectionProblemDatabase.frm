VERSION 5.00
Begin VB.Form frmConnectionProblemDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "From Connection Problem Database"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSimpan 
      Caption         =   "Simpan dan hubungkan"
      Height          =   510
      Left            =   3060
      TabIndex        =   9
      Top             =   1665
      Width           =   1860
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "Batal"
      Height          =   510
      Left            =   4995
      TabIndex        =   8
      Top             =   1665
      Width           =   1230
   End
   Begin VB.TextBox txtDatabse 
      Height          =   330
      Left            =   1215
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1260
      Width           =   5010
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      Left            =   1215
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   855
      Width           =   5010
   End
   Begin VB.TextBox txtUsername 
      Height          =   330
      Left            =   1215
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   450
      Width           =   5010
   End
   Begin VB.TextBox txtHost 
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   45
      Width           =   5010
   End
   Begin VB.Label Label4 
      Caption         =   "Database"
      Height          =   330
      Left            =   45
      TabIndex        =   6
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   855
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Host"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1095
   End
End
Attribute VB_Name = "frmConnectionProblemdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBatal_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    Dim dataFile() As String
    Dim i As Integer

    Dim file As String

    file = "\connectionconf.dat"

    Dim host As String
    Dim username As String
    Dim password As String
    Dim database As String

    Dim data() As String

    ReDim Preserve data(0)
    data(0) = "host:" & txtHost.Text
    ReDim Preserve data(1)
    data(1) = "username:" & txtUsername.Text
    ReDim Preserve data(2)
    data(2) = "password:" & txtPassword.Text
    ReDim Preserve data(3)
    data(3) = "database:" & txtDatabse.Text

    createAndSaveFile file, data


    If konekToServer Then
        MsgBox "koneksi sukses"
        frmMenuUtama.Show
        Unload Me
    Else
        MsgBox "koneksi gagal"
    End If

End Sub

Private Sub Form_Load()
    Clear Me

    Dim dataFile() As String
    Dim i As Integer

    Dim file As String

    file = "\connectionconf.dat"

    dataFile = loadDataFromFile(file)

    txtHost.Text = getDataFromArray(dataFile, "host")
    txtUsername.Text = getDataFromArray(dataFile, "username")
    txtPassword.Text = getDataFromArray(dataFile, "password")
    txtDatabse.Text = getDataFromArray(dataFile, "database")
End Sub
