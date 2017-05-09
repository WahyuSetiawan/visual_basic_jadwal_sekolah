VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   1380
      TabIndex        =   5
      Top             =   900
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   400
      Left            =   90
      TabIndex        =   4
      Top             =   900
      Width           =   1245
   End
   Begin VB.TextBox txt_password 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "x"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   420
      Width           =   1605
   End
   Begin VB.TextBox txt_username 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   30
      Width           =   1605
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Login
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    Clear Me
End Sub

Private Sub Form_Load()
    Clear Me
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Me.Login
End Sub

Sub Login()
    Dim username As String
    Dim password As String
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select * from operator where username = '" & Trim$(Me.txt_username.Text) & "' and pass = '" & Trim$(Me.txt_password.Text) & "'")

    If Not rs.EOF Then
        Load frmMenuUtama
        frmMenuUtama.Show
        Me.Visible = False
    Else
        MsgBox "Password atau username salah", vbOKOnly, "P E R H A T I A N"
    End If

    Call closeRecordset(rs)

End Sub

Private Sub txt_username_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txt_password.SetFocus
End Sub
