VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClientUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Client"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14850
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameTelahLogin 
      Height          =   5775
      Left            =   2970
      TabIndex        =   14
      Top             =   810
      Width           =   9015
      Begin VB.CommandButton cmdPelajaranSelesai 
         Caption         =   "Pelajaran Selesai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2970
         TabIndex        =   29
         Top             =   4545
         Width           =   3120
      End
      Begin VB.Label lblNamaPelajaran 
         Caption         =   "Pelajaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3465
         TabIndex        =   28
         Top             =   2295
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "Nama Pelajaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         TabIndex        =   27
         Top             =   2295
         Width           =   2850
      End
      Begin VB.Label NIP 
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         TabIndex        =   26
         Top             =   1800
         Width           =   2220
      End
      Begin VB.Label lblSisaWaktu 
         Caption         =   "Sisa Waktu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3465
         TabIndex        =   25
         Top             =   3825
         Width           =   1995
      End
      Begin VB.Label Label10 
         Caption         =   "Sisa Waktu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         TabIndex        =   24
         Top             =   3825
         Width           =   2220
      End
      Begin VB.Label lblJamSelesai 
         Caption         =   "Jam Selesai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3465
         TabIndex        =   23
         Top             =   3285
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Jam Selesai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         TabIndex        =   22
         Top             =   3285
         Width           =   2220
      End
      Begin VB.Label lblJamMulai 
         Caption         =   "Jam Mulai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3465
         TabIndex        =   21
         Top             =   2790
         Width           =   1995
      End
      Begin VB.Label Label6 
         Caption         =   "Jam Mulai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         TabIndex        =   20
         Top             =   2790
         Width           =   2220
      End
      Begin VB.Label lblNIP 
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3465
         TabIndex        =   19
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lblNamaGuru 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Guru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   0
         TabIndex        =   15
         Top             =   810
         Width           =   8970
      End
   End
   Begin VB.Timer Timerstatus 
      Interval        =   2
      Left            =   1395
      Top             =   0
   End
   Begin VB.Timer TimerRequest 
      Interval        =   2000
      Left            =   900
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   7500
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   200
      Left            =   450
      Top             =   0
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "Option"
      Height          =   465
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1185
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Framelogin 
      Height          =   3500
      Left            =   2700
      TabIndex        =   7
      Top             =   1080
      Width           =   9000
      Begin VB.TextBox txtNipGuru 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1260
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2430
         Width           =   6675
      End
      Begin VB.Label lblWaktuSekarang 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3885
         TabIndex        =   18
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label lblPelajaran 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1260
         TabIndex        =   17
         Top             =   1575
         Width           =   3375
      End
      Begin VB.Label lblGuru 
         Alignment       =   1  'Right Justify
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   5220
         TabIndex        =   16
         Top             =   1530
         Width           =   2715
      End
   End
   Begin VB.Frame FrameOption 
      Height          =   3500
      Left            =   6840
      TabIndex        =   1
      Top             =   1980
      Visible         =   0   'False
      Width           =   9000
      Begin VB.CommandButton cmdDataKelas 
         Caption         =   "..."
         Height          =   375
         Left            =   5490
         TabIndex        =   13
         Top             =   1665
         Width           =   375
      End
      Begin VB.TextBox txtKelas 
         Height          =   375
         Left            =   2205
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1665
         Width           =   3300
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "Simpan"
         Height          =   510
         Left            =   6390
         TabIndex        =   10
         Top             =   1080
         Width           =   2355
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   510
         Left            =   6390
         TabIndex        =   6
         Top             =   450
         Width           =   2355
      End
      Begin VB.TextBox txtPort 
         Height          =   420
         Left            =   2205
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1125
         Width           =   3300
      End
      Begin VB.TextBox txtIP 
         Height          =   420
         Left            =   2205
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   495
         Width           =   3300
      End
      Begin VB.Label Label3 
         Caption         =   "Kelas"
         Height          =   420
         Left            =   270
         TabIndex        =   11
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         Height          =   375
         Left            =   315
         TabIndex        =   5
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "Client.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "frmClientUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const delimiterBaris As String = "|"
Const delimiterKolom As String = "#"
Const delimiterData As String = "$"

Public kelas As String
Public pelajaran As String
Public namaGuru As String

Public idKelas As Integer
'Public NIP As String
Public idGuru As String
Public idPelajaran As String
Public idJadwal As String
Public idRekapJadwal As String

Public jamMulai As Double
Public jamSelesai As Double

Public pelajaranAktif As Boolean

Public time As Date

Function pelajaranSelesai()
    frmClientUtama.SetFocus

    MsgBox "Waktu Pelajaran Berakhir"

    Me.pelajaranAktif = Not Me.pelajaranAktif

    Dim Datasend As String
    Dim dataTmp1(1) As String

    dataTmp1(0) = delimiterData & "logoutGuru"
    dataTmp1(1) = CStr(idRekapJadwal)

    Datasend = Join(dataTmp1, delimiterBaris)

    wskClient.SendData Datasend

    Me.Framelogin.Visible = True
    Me.FrameOption.Visible = False
    Me.FrameTelahLogin.Visible = False

    Me.idKelas = 0
    Me.idGuru = ""
    Me.idPelajaran = ""
    Me.idJadwal = ""
    Me.idRekapJadwal = ""
    Me.lblNamaGuru.Caption = ""
    Me.lblNamaPelajaran.Caption = ""

    lblPelajaran.Caption = ""
    lblGuru.Caption = ""
End Function

Private Sub cmdConnect_Click()
    If txtIP.Text = "" Then txtIP.SetFocus
    If txtPort.Text = "" Then txtPort.SetFocus

    If wskClient.State = sckConnecting Then wskClient.Close
    If wskClient.State = sckConnected Then wskClient.Close
    wskClient.Connect txtIP.Text, txtPort.Text

    If wskClient.State = sckConnected Then wskClient.SendData CStr("daftarkelas" & delimiterBaris & CStr(idKelas))
End Sub

Private Sub cmdDataKelas_Click()
    frmDataKelas.Show , Me
End Sub

Private Sub cmdOption_Click()
    If FrameOption.Visible Then
        If Not pelajaranAktif Then
            Framelogin.Visible = True
        Else
            FrameTelahLogin.Visible = True
        End If

        FrameOption.Visible = False
    Else
        FrameTelahLogin.Visible = False
        Framelogin.Visible = False
        FrameOption.Visible = True
    End If
End Sub

Private Sub cmdPelajaranSelesai_Click()
    pelajaranSelesai
End Sub

Private Sub CmdSimpan_Click()
    Dim dataConf() As String

    ReDim Preserve dataConf(0)
    dataConf(0) = "ip:" + txtIP.Text
    ReDim Preserve dataConf(1)
    dataConf(1) = "port:" + txtPort.Text
    ReDim Preserve dataConf(2)
    dataConf(2) = "idkelas:" + CStr(idKelas)

    createAndSaveFile pathFileConfiguration, dataConf
End Sub

Private Sub Form_Load()
    lblPelajaran.Caption = ""
    lblGuru.Caption = ""

    TimerWaktu.Enabled = True

    Clear Me

    Me.Framelogin.Visible = True
    Me.FrameOption.Visible = False
    Me.FrameTelahLogin.Visible = False

    Dim strFile As String
    Dim intFile As Integer

    txtIP.Text = getDataFromFile(pathFileConfiguration, "ip")
    txtPort.Text = getDataFromFile(pathFileConfiguration, "port")
    If Not getDataFromFile(pathFileConfiguration, "idkelas") = "" Then
        idKelas = CInt(getDataFromFile(pathFileConfiguration, "idkelas"))
    End If

    dataSplit = Split(getDataFromFile("\dataTMP.dat", "datakelas"), delimiterBaris)

    For i = LBound(dataSplit) To UBound(dataSplit)

        If Not dataSplit(i) = "" Then
            dataKolom = Split(dataSplit(i), delimiterKolom)

            If dataKolom(0) = CStr(idKelas) Then
                txtKelas.Text = dataKolom(0) + " - " + dataKolom(1)
            End If
        End If
    Next i
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Framelogin.Left = (Me.Width / 2) - (Me.Framelogin.Width / 2)
    Framelogin.Top = (Me.Height / 2) - (Me.Framelogin.Height / 2)

    FrameOption.Left = (Me.Width / 2) - (Me.FrameOption.Width / 2)
    FrameOption.Top = (Me.Height / 2) - (Me.FrameOption.Height / 2)

    FrameTelahLogin.Left = (Me.Width / 2) - (Me.FrameTelahLogin.Width / 2)
    FrameTelahLogin.Top = (Me.Height / 2) - (Me.FrameTelahLogin.Height / 2)

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

Private Sub TimerRequest_Timer()
    If Not pelajaranAktif Then
        If wskClient.State = sckConnected Then
            getJadwal
        End If
    Else
        If (jamSelesai <= dateToSeconds(Now)) Then
            pelajaranSelesai

        Else
            lblJamMulai.Caption = secondToDate(jamMulai)
            lblJamSelesai.Caption = secondToDate(jamSelesai)

            lblSisaWaktu.Caption = secondToDate(jamSelesai - dateToSeconds(Now))
        End If

    End If

    If Not Me.pelajaranAktif Then
        lblJamMulai.Caption = secondToDate(jamMulai)
        lblJamSelesai.Caption = secondToDate(jamSelesai)
    End If
End Sub

Private Sub Timerstatus_Timer()
    Dim pesan As String

    Select Case wskClient.State

        Case sckClosing
            pesan = "Tidak terhubung dengan server"
            MsgBox pesan

            wskClient.Close
        Case sckOpen
            pesan = "Saluran Terbuka"
        Case sckListening
            pesan = "Saluran Menunggu"
        Case sckConnectionPending
            pesan = "Saluran terganggu"
        Case sckResolvingHost
            pesan = "Saluran menghilang"
        Case sckHostResolved
            pesan = "Saluran server tidak ditemukan"
        Case sckConnecting
            pesan = "Saluran menhubungkan"
        Case sckConnected
            pesan = "Saluran Tehubung"
        Case sckClosing
            pesan = "Saluran Menutup"
        Case sckError
            pesan = "Saluran terganggu"

            wskClient.Close
    End Select


    StatusBar.Panels.Item(1).Text = pesan
    StatusBar.Panels.Item(2).Text = "Pelajaran:" & CStr(Me.idPelajaran)
    StatusBar.Panels.Item(3).Text = "Guru:" & CStr(Me.idGuru)
    StatusBar.Panels.Item(4).Text = "Jam Mulai:" & secondToDate(jamMulai)
    StatusBar.Panels.Item(5).Text = "Jam Selesai:" & secondToDate(jamSelesai)
    StatusBar.Panels.Item(6).Text = "Rekap Jadwal:" & Me.idRekapJadwal
    StatusBar.Panels.Item(7).Text = Me.pelajaranAktif
    StatusBar.Panels.Item(8).Text = Me.idJadwal
End Sub

Private Sub TimerWaktu_Timer()
    lblWaktuSekarang.Caption = CStr(Format$(Now, "hh:mm:ss"))

    If pelajaranAktif Then

    End If
End Sub

Private Sub txtNipGuru_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then Login txtNipGuru.Text
End Sub

Private Sub wskClient_Connect()
    wskClient.SendData delimiterData & "getDataKelas"
End Sub

Private Sub getJadwal()
    idKelas = getDataFromFile(pathFileConfiguration, "idkelas")

    If wskClient.State = sckConnected Then
        Dim satasend As String
        Dim dataTmp(1) As String

        dataTmp(0) = delimiterData & "getJadwal"
        dataTmp(1) = CStr(idKelas)

        Datasend = Join(dataTmp, delimiterBaris)

        wskClient.SendData Datasend
    End If
End Sub

Private Sub Login(NIP As String)
    If wskClient.State = sckConnected Then
        Dim Datasend As String
        Dim dataTmp1(3) As String

        dataTmp1(0) = delimiterData & "loginGuru"
        dataTmp1(1) = CStr(idGuru)
        dataTmp1(2) = Trim$(CStr(NIP))
        dataTmp1(3) = CStr(idRekapJadwal)

        Datasend = Join(dataTmp1, delimiterBaris)

        wskClient.SendData Datasend
    End If
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim datakelas() As String
    Dim dataSplit() As String

    ReDim Preserve datakelas(0)

    Dim dataKolom() As String
    Dim dataKolom2() As String
    Dim dataTmp As String

    wskClient.GetData dataarrived, vbString

    dataSplit = Split(dataarrived, delimiterBaris)

    If dataSplit(0) = "dataKelas" Then
        If dataSplit(1) = "EOF" Then
            'MsgBox "Tidak Terdapat data kelas di dalam server"
        Else
            dataSplit = Split(dataarrived, delimiterBaris)

            For i = LBound(dataSplit) + 1 To UBound(dataSplit)
                If Not dataSplit(i) = "" Then
                    dataTmp = dataTmp + dataSplit(i) + delimiterBaris
                End If
            Next i

            datakelas(0) = "datakelas:" + dataTmp

            createAndSaveFile "\dataTMP.dat", datakelas
        End If
    ElseIf dataSplit(0) = "dataJadwal" Then
        If dataSplit(1) = "EOF" Then

            lblPelajaran.Caption = ""
            lblGuru.Caption = ""
            lblNamaPelajaran.Caption = ""
            lblNamaGuru.Caption = ""
            lblJamMulai.Caption = ""
            lblJamSelesai.Caption = ""

            Me.idPelajaran = ""
            Me.idGuru = ""
            jamMulai = 0
            Me.jamSelesai = 0
            Me.idJadwal = ""

            lblGuru.Caption = ""
            lblPelajaran.Caption = ""
        Else
            dataKolom = Split(dataSplit(1), delimiterKolom)
            dataKolom2 = Split(dataSplit(2), delimiterKolom)

            Me.idPelajaran = dataKolom(7)
            Me.idGuru = dataKolom(8)
          '  Me.jamMulai = CDate(CStr(dataKolom(3) \ 60) & ":" & CStr(dataKolom(3) - ((dataKolom(3) \ 60) * 60)))
            Me.jamMulai = CDbl(dataKolom(3))
            Me.jamSelesai = CDbl(dataKolom(4))
          '  Me.jamSelesai = CDate(CStr(dataKolom(4) \ 60) & ":" & CStr(dataKolom(4) - ((dataKolom(4) \ 60) * 60)))
            Me.idJadwal = CStr(dataKolom(0))
            Me.idRekapJadwal = dataKolom2(0)

            If Not dataKolom(7) = "" Then
                Dim Datasend As String
                Dim dataTmp1(2) As String

                dataTmp1(0) = delimiterData & "getDataPelajaranAndGuru"
                dataTmp1(1) = CStr(idPelajaran)
                dataTmp1(2) = CStr(idGuru)

                Datasend = Join(dataTmp1, delimiterBaris)
            End If

            wskClient.SendData Datasend
        End If
    ElseIf dataSplit(0) = "LoginCheck" Then
        If dataSplit(1) = "EOF" Then
            MsgBox "login tidak berhasil"
        Else
            dataKolom = Split(dataSplit(1), delimiterKolom)
            dataKolom2 = Split(dataSplit(2), delimiterKolom)

            Me.idRekapJadwal = dataKolom2(0)

            Me.lblNamaGuru.Caption = dataKolom(1)
            Me.lblNamaGuru.Caption = dataKolom(1)
            Me.lblNIP.Caption = dataKolom(3)

            MsgBox "Login Berhasil"

            Me.pelajaranAktif = True

            Me.Framelogin.Visible = False
            Me.FrameOption.Visible = False
            Me.FrameTelahLogin.Visible = True
        End If
    ElseIf dataSplit(0) = "dataPelajaranDanGuru" Then
        If dataSplit(1) = "EOF" Then
            'MsgBox "Tidak Terdapat data kelas di dalam server"
        Else
            lblPelajaran.Caption = dataSplit(1)
            lblNamaPelajaran.Caption = dataSplit(1)
            lblGuru.Caption = dataSplit(2)
            lblNamaGuru.Caption = dataSplit(2)
        End If
    End If
End Sub



