VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenuUtama 
   Caption         =   "Menu Utama"
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start Server"
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   375
      Width           =   1335
   End
   Begin VB.TextBox txtlog 
      Height          =   6915
      Left            =   90
      TabIndex        =   1
      Text            =   "txtlog"
      Top             =   1350
      Width           =   5640
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   9465
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4590
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7799
   End
   Begin VB.Label lblPortServer 
      AutoSize        =   -1  'True
      Caption         =   "Port Server :"
      Height          =   195
      Left            =   3060
      TabIndex        =   8
      Top             =   855
      Width           =   885
   End
   Begin VB.Label lblIPServer 
      AutoSize        =   -1  'True
      Caption         =   "IP Server:"
      Height          =   195
      Left            =   3060
      TabIndex        =   7
      Top             =   495
      Width           =   705
   End
   Begin VB.Label lblStatusServer 
      AutoSize        =   -1  'True
      Caption         =   "Server Status"
      Height          =   195
      Left            =   3060
      TabIndex        =   6
      Top             =   135
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port Server :"
      Height          =   195
      Left            =   1980
      TabIndex        =   5
      Top             =   855
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP Server :"
      Height          =   195
      Left            =   1980
      TabIndex        =   4
      Top             =   495
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Status :"
      Height          =   195
      Left            =   1980
      TabIndex        =   3
      Top             =   135
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "Server.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "Program"
      Begin VB.Menu mnuPeraturan 
         Caption         =   "Peraturan"
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Server"
      End
      Begin VB.Menu mnukeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "Data"
      Begin VB.Menu mnudataguru 
         Caption         =   "Data Guru"
      End
      Begin VB.Menu mnudatapelajaran 
         Caption         =   "Data Pelajaran"
      End
      Begin VB.Menu mnuPengaturanDataInputan 
         Caption         =   "Pengaturan Data Inputan"
      End
      Begin VB.Menu mnudatajadwal 
         Caption         =   "Data Jadwal"
      End
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indexstatusbarserver As Integer


Sub ListMataPelajaran()
With Me.lstMataPelajaran

.View = lvwReport
.FullRowSelect = True
.GridLines = True
.AllowColumnReorder = False

.ColumnHeaders.Clear
.ColumnHeaders.Add , , "No", 500
.ColumnHeaders.Add , , "ID Anggota", 2000
.ColumnHeaders.Add , , "Nama Anggota", 2000
.ColumnHeaders.Add , , "Alamat", 1000
.ColumnHeaders.Add , , "Telepon", 800
.ColumnHeaders.Add , , "No Rekening", 800

KONEKSI

rs.CursorLocation = adUseClient
rs.Open "select * from pelajaran", Conn, adOpenDynamic, adLockOptimistic

.ListItems.Clear

Dim i As Integer
i = 1

If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        Set List = .ListItems.Add(, , i)
        
        If IsNull(rs.Fields("id_anggota").value) = True Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("id_anggota")
        End If
            
        If IsNull(rs.Fields("nama_anggota").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("nama_anggota")
        End If
        
        If IsNull(rs.Fields("alamat").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("alamat")
        End If
        
        If IsNull(rs.Fields("telp").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("telp")
        End If
        
        If IsNull(rs.Fields("no_rekening").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("no_rekening")
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
End If

Conn.Close

Set Conn = Nothing
Set rs = Nothing

End With
End Sub

Sub ListJadwal()
With Me.lstJadwal

.View = lvwReport
.FullRowSelect = True
.GridLines = True
.AllowColumnReorder = False

.ColumnHeaders.Clear
.ColumnHeaders.Add , , "No", 500
.ColumnHeaders.Add , , "ID Anggota", 2000
.ColumnHeaders.Add , , "Nama Anggota", 2000
.ColumnHeaders.Add , , "Alamat", 1000
.ColumnHeaders.Add , , "Telepon", 800
.ColumnHeaders.Add , , "No Rekening", 800

KONEKSI

rs.CursorLocation = adUseClient
rs.Open "select * from jadwal", Conn, adOpenDynamic, adLockOptimistic

.ListItems.Clear

Dim i As Integer
i = 1

If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        Set List = .ListItems.Add(, , i)
        
        If IsNull(rs.Fields("id_anggota").value) = True Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("id_anggota")
        End If
            
        If IsNull(rs.Fields("nama_anggota").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("nama_anggota")
        End If
        
        If IsNull(rs.Fields("alamat").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("alamat")
        End If
        
        If IsNull(rs.Fields("telp").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("telp")
        End If
        
        If IsNull(rs.Fields("no_rekening").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("no_rekening")
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
End If

Conn.Close

Set Conn = Nothing
Set rs = Nothing

End With
End Sub

Sub ListWaktuPelajaran()
With Me.lstJamPelajaran

.View = lvwReport
.FullRowSelect = True
.GridLines = True
.AllowColumnReorder = False

.ColumnHeaders.Clear
.ColumnHeaders.Add , , "No", 500
.ColumnHeaders.Add , , "ID Anggota", 2000
.ColumnHeaders.Add , , "Nama Anggota", 2000
.ColumnHeaders.Add , , "Alamat", 1000
.ColumnHeaders.Add , , "Telepon", 800
.ColumnHeaders.Add , , "No Rekening", 800

KONEKSI

rs.CursorLocation = adUseClient
rs.Open "select * from jadwal", Conn, adOpenDynamic, adLockOptimistic

.ListItems.Clear

Dim i As Integer
i = 1

If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        Set List = .ListItems.Add(, , i)
        
        If IsNull(rs.Fields("id_anggota").value) = True Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("id_anggota")
        End If
            
        If IsNull(rs.Fields("nama_anggota").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("nama_anggota")
        End If
        
        If IsNull(rs.Fields("alamat").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("alamat")
        End If
        
        If IsNull(rs.Fields("telp").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("telp")
        End If
        
        If IsNull(rs.Fields("no_rekening").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("no_rekening")
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
End If

Conn.Close

Set Conn = Nothing
Set rs = Nothing

End With
End Sub

Sub ListKelas()
With Me.lstKelas

.View = lvwReport
.FullRowSelect = True
.GridLines = True
.AllowColumnReorder = False

.ColumnHeaders.Clear
.ColumnHeaders.Add , , "No", 500
.ColumnHeaders.Add , , "Nama", 2000
.ColumnHeaders.Add , , "", 2000
.ColumnHeaders.Add , , "Alamat", 1000
.ColumnHeaders.Add , , "Telepon", 800
.ColumnHeaders.Add , , "No Rekening", 800

KONEKSI

rs.CursorLocation = adUseClient
rs.Open "select * from kelas", Conn, adOpenDynamic, adLockOptimistic

.ListItems.Clear

Dim i As Integer
i = 1

If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        Set List = .ListItems.Add(, , i)
        
        If IsNull(rs.Fields("id_anggota").value) = True Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("id_anggota")
        End If
            
        If IsNull(rs.Fields("nama_anggota").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("nama_anggota")
        End If
        
        If IsNull(rs.Fields("alamat").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("alamat")
        End If
        
        If IsNull(rs.Fields("telp").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("telp")
        End If
        
        If IsNull(rs.Fields("no_rekening").value) Then
            List.ListSubItems.Add , , ""
        Else
            List.ListSubItems.Add , , rs.Fields("no_rekening")
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
End If

Conn.Close

Set Conn = Nothing
Set rs = Nothing

End With
End Sub

Private Sub cmdStartServer_Click()
If Winsock1.State <> 1 Then Winsock1.Close

Winsock1.Listen

StatusBar.Panels.Item(1).Text = "Server Nyala"

If Winsock1.State = 2 Then
    lblIPServer.Caption = Winsock1.LocalIP
    lblPortServer.Caption = Winsock1.LocalPort
    lblStatusServer.Caption = "Server nyala"
Else
    lblStatusServer.Caption = "Server tidak dapat dinyalakan"
End If

End Sub


Private Sub Form_Load()
Clear Me
End Sub

Private Sub Form_Resize()

On Error Resume Next

Dim minFormWidth As Integer
Dim minFormHeight As Integer

'ubah data form min yang akan dipakai
minFormWidth = 6000
minFormHeight = 6000

If MaxMinForm(Me, minFormHeight, minFormWidth) Then
    'masukan datanya disini yang lain dari ini hanyalah template
    FitInScreen Me, Me.Image1
End If

If Me.ScaleHeight < minFormHeight Then
    HoldFormScaleHeight Me, minFormHeight
End If
    
If Me.ScaleWidth < minFormWidth Then
    HoldFormScaleWidth Me, minFormWidth
End If

End Sub

Private Sub mnudataguru_Click()
frmListDataGuru.Show , Me
End Sub

Private Sub mnukeluar_Click()
End
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData dataArrived, vbString


txtlog.Text = txtlog.Text & dataArrived & "/n"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Winsock1.State <> 1 Then Winsock1.Close

lblStatusServer.Caption = "Server Error"
End Sub
