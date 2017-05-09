VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenuUtama 
   Caption         =   "Menu Utama"
   ClientHeight    =   9750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerStatus 
      Interval        =   200
      Left            =   5040
      Top             =   630
   End
   Begin VB.ComboBox cmbTahun 
      Height          =   315
      Left            =   6615
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   495
      Width           =   1905
   End
   Begin VB.ComboBox cmbSemester 
      Height          =   315
      Left            =   6615
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   90
      Width           =   1905
   End
   Begin VB.CommandButton cmdRefreshJadwal 
      Caption         =   "Clear Log"
      Height          =   465
      Left            =   5850
      TabIndex        =   12
      Top             =   810
      Width           =   1725
   End
   Begin VB.ListBox ListLogServer 
      Height          =   7665
      Left            =   180
      TabIndex        =   11
      Top             =   1440
      Width           =   8340
   End
   Begin VB.Timer TimerSessionManager 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4590
      Top             =   630
   End
   Begin VB.CommandButton cmdPengaturan 
      Caption         =   "Pengaturan"
      Height          =   420
      Left            =   180
      TabIndex        =   8
      Top             =   765
      Width           =   1320
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start Server"
      Height          =   495
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   9465
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   4590
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7799
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Semester :"
      Height          =   195
      Left            =   5805
      TabIndex        =   10
      Top             =   180
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tahun :"
      Height          =   195
      Left            =   5805
      TabIndex        =   9
      Top             =   540
      Width           =   555
   End
   Begin VB.Label lblPortServer 
      AutoSize        =   -1  'True
      Caption         =   "Port Server :"
      Height          =   195
      Left            =   3060
      TabIndex        =   7
      Top             =   855
      Width           =   885
   End
   Begin VB.Label lblIPServer 
      AutoSize        =   -1  'True
      Caption         =   "IP Server:"
      Height          =   195
      Left            =   3060
      TabIndex        =   6
      Top             =   495
      Width           =   705
   End
   Begin VB.Label lblStatusServer 
      AutoSize        =   -1  'True
      Caption         =   "Server Status"
      Height          =   195
      Left            =   3060
      TabIndex        =   5
      Top             =   135
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port Server :"
      Height          =   195
      Left            =   1980
      TabIndex        =   4
      Top             =   855
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP Server :"
      Height          =   195
      Left            =   1980
      TabIndex        =   3
      Top             =   495
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Status :"
      Height          =   195
      Left            =   1980
      TabIndex        =   2
      Top             =   135
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "frmMenuUtama.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "Program"
      Begin VB.Menu mnuPeraturan 
         Caption         =   "Pengaturan"
      End
      Begin VB.Menu mnuOperator 
         Caption         =   "Operator"
      End
      Begin VB.Menu mnukeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu MnuDataGuru 
         Caption         =   "Data Guru"
      End
      Begin VB.Menu mnudatapelajaran 
         Caption         =   "Data Pelajaran"
      End
      Begin VB.Menu mnuDataSemester 
         Caption         =   "Data Semester"
      End
      Begin VB.Menu mnuDataKelas 
         Caption         =   "Data Kelas"
      End
      Begin VB.Menu mnudatajadwal 
         Caption         =   "Data Jadwal"
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu MenuLaporanPelajaran 
         Caption         =   "Laporan Pelajaran"
      End
      Begin VB.Menu mnuLaporaDataKelas 
         Caption         =   "Laporan Data Kelas"
      End
      Begin VB.Menu MnuLaporanGuru 
         Caption         =   "Laporan Data Guru"
      End
      Begin VB.Menu mnuJadwalPelajaranGuru 
         Caption         =   "Laporan Jadwal Pelajaran Guru"
      End
      Begin VB.Menu mnuJadwalPelajaranKelas 
         Caption         =   "Laporan Jadwal Pelajaran Kelas"
      End
      Begin VB.Menu menuPelajaranDanGuru 
         Caption         =   "Laporan Daftar Pelajaran Dan Guru"
      End
      Begin VB.Menu mnuLaporanRekap 
         Caption         =   "Laporan Rekap Presensi"
      End
   End
End
Attribute VB_Name = "frmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indexstatusbarserver As Integer

Const delimiterBaris As String = "|"
Const delimiterKolom As String = "#"
Const delimiterData As String = "$"

Public semester As String
Public tahun As String
Public port As String

Private Sub cmbTahun_Change()
    Me.tahun = cmbTahun.Text
End Sub

Private Sub cmdPengaturan_Click()
    frmPengaturan.Show
End Sub

Private Sub cmdRefreshJadwal_Click()
    Me.ListLogServer.Clear
End Sub

Private Sub cmdStartServer_Click()
    StartServer 0
End Sub

Private Sub Form_Activate()
    Dim dataFile() As String
    Dim i As Integer

    dataFile = loadDataFromFile("\conf.dat")

    Me.semester = getDataFromArray(dataFile, "idsemester")
    Me.tahun = getDataFromArray(dataFile, "tahun")
    Me.port = getDataFromArray(dataFile, "port")
End Sub

Private Sub Form_Load()
    Clear Me

    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select distinct(tahun) from jadwal  where deleted  = 0")

    cmbTahun.Clear

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cmbTahun.AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)

    Set rs = openRecordset("select * from semester where deleted  = 0")

    cmbSemester.Clear

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            cmbSemester.AddItem rs.Fields(1)
            rs.MoveNext
        Loop
    End If

    Call closeRecordset(rs)

    Dim dataFile() As String
    Dim i As Integer

    dataFile = loadDataFromFile("\conf.dat")

    cmbSemester.Text = getDataFromArray(dataFile, "semester")
    cmbTahun.Text = getDataFromArray(dataFile, "tahun")
    wskServer(0).LocalPort = getDataFromArray(dataFile, "port")

    Set rs = openRecordset("select * from semester where nama = '" & Trim$(cmbSemester.Text) & "' and deleted  = 0")

    If Not rs.EOF Then
        Me.semester = rs.Fields(0)
    End If

    Call closeRecordset(rs)
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

Private Sub lblTahunAktif_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogin.Visible = True
End Sub

Private Sub MenuLaporanPelajaran_Click()
    frmLaporanDataPelajaran.Show , Me
End Sub

Private Sub menuPelajaranDanGuru_Click()
    frmLaporanPelajaranDanGruu.Show , Me
End Sub

Private Sub mnudataguru_Click()
    frmListDataGuru.Show , Me
End Sub

Private Sub mnudatajadwal_Click()
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select count(*) as jumlah from semester where deleted = 0")

    If Not rs.EOF Then
        If rs.Fields("jumlah") = 0 Then
            MsgBox "Belum ada data tentang semester, Silahkan anda mengisi data semester terlebih dahulu"
            frmListDataSemester.Show , Me
            Exit Sub
        End If
    End If

    Call closeRecordset(rs)


    Set rs = openRecordset("select count(*) as jumlah from kelas where deleted = 0")

    If Not rs.EOF Then
        If rs.Fields("jumlah") = 0 Then
            MsgBox "Belum ada data tentang kelas, Silahkan anda mengisi data semester terlebih dahulu"
            frmListDataKelas.Show , Me
            Exit Sub
        End If
    End If

    Call closeRecordset(rs)


    Set rs = openRecordset("select count(*) as jumlah from guru where deleted = 0")

    If Not rs.EOF Then
        If rs.Fields("jumlah") = 0 Then
            MsgBox "Belum ada data tentang guru, Silahkan anda mengisi data semester terlebih dahulu"
            frmListDataGuru.Show , Me
            Exit Sub
        End If
    End If

    Call closeRecordset(rs)

    Set rs = openRecordset("select count(*) as jumlah from pelajaran where deleted = 0")

    If Not rs.EOF Then
        If rs.Fields("jumlah") = 0 Then
            MsgBox "Belum ada data tentang pelajaran, Silahkan anda mengisi data semester terlebih dahulu"
            frmListMataPelajaran.Show , Me
            Exit Sub
        End If
    End If

    Call closeRecordset(rs)

    frmListJadwal.Show , Me
End Sub

Private Sub mnuDataKelas_Click()
    frmListDataKelas.Show , Me
End Sub

Private Sub mnudatapelajaran_Click()
    frmListMataPelajaran.Show , Me
End Sub

Private Sub mnuDataSemester_Click()
    frmListDataSemester.Show , Me
End Sub

Private Sub mnuJadwalPelajaranGuru_Click()
    frmLaporanPelajaranGuru.Show , Me
End Sub

Private Sub mnuJadwalPelajaranKelas_Click()
    frmLaporanPelajaranKelas.Show , Me
End Sub

Private Sub mnukeluar_Click()
    End
End Sub

Private Sub mnuLaporaDataKelas_Click()
    frmLaporanDataKelas.Show , Me
End Sub

Private Sub MnuLaporanGuru_Click()
    FrmLaporanDataGuru.Show , Me
End Sub

Private Sub mnuLaporanRekap_Click()
    frmLaporanRekapPresensi.Show , Me
End Sub

Private Sub mnuOperator_Click()
    frmOperator.Show , Me
End Sub

Private Sub mnuPeraturan_Click()
    frmPengaturan.Show , Me
End Sub

Sub StartServer(Index As Integer)
    If wskServer(Index).State = sckListening Then wskServer(Index).Close

    If Not getDataFromFile(pathFileConfiguration, "port") = "" Then wskServer(Index).LocalPort = CLng(getDataFromFile(pathFileConfiguration, "port"))

    wskServer(Index).Listen

    StatusBar.Panels.Item(1).Text = "Server Nyala"

    If wskServer(Index).State = 2 Then
        lblIPServer.Caption = wskServer(Index).LocalIP
        lblPortServer.Caption = wskServer(Index).LocalPort
        lblStatusServer.Caption = "Server nyala"
        Me.TimerSessionManager.Enabled = True
        ListLogServer.AddItem TimeValue(Now) & " : Server dimulai Dengan ip : " & wskServer(Index).LocalIP & " dan port : " & wskServer(Index).LocalPort, 0
        ListLogServer.AddItem TimeValue(Now) & " : Dengan server di semester : " & cmbSemester.Text & " dan id : " & semester & " serta tahun : " & tahun, 0
    Else
        lblStatusServer.Caption = "Server tidak dapat dinyalakan"
        ListLogServer.AddItem TimeValue(Now) & " : Tidak dapat menjalankan Server", 0
        TimerBroadcastMessage.Enabled = False
    End If
End Sub

Sub ServerClose(Index As Integer)
    If wskServer(Index).State <> sckClosed Then wskServer(Index).Close

    ListLogServer.AddItem TimeValue(Now) & " : Client : " & Index & " Terputus", 0
End Sub

Sub ServerConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Long
    Dim j As Long

    On Error GoTo errHandle

    If Index = 0 Then
        For i = 0 To wskServer.UBound
            If wskServer(i).State = sckClosed Or wskServer(i).State = sckClosing Then
                j = i
                Exit For
            End If
        Next i

        If j = 0 Then
            Call Load(wskServer(wskServer.UBound + 1))
            j = wskServer.UBound
            ListLogServer.AddItem TimeValue(Now) & " : Server Baru Terbuat Server = " & j, 0
        End If

        wskServer(j).Close
        wskServer(j).Accept requestID

        ListLogServer.AddItem TimeValue(Now) & " : Server " & j & " Telah Tersambung", 0
    End If

    Exit Sub

errHandle:
    Call wskServer(0).Close
End Sub
Sub ServerDataArrived(Index As Integer)
    'On Error Resume Next
    
    Dim dataSplit() As String
    Dim i As Integer
    Dim query As String

    Dim BarisTMP() As String
    Dim DataTMP() As String
    Dim KolomTMP() As String

    Dim sendTMP() As String

    Dim dateSend As String
    Dim log As String

    Dim rs As ADODB.Recordset

                Dim rsDB As ADODB.Recordset

    wskServer(Index).GetData dataArrived, vbString

    'ListLogServer.AddItem TimeValue(Now) & " : Pesan yang diterima " & dataArrived, 0

    DataTMP = Split(dataArrived, delimiterData)


    For j = LBound(DataTMP) To UBound(DataTMP)
        If Not DataTMP(j) = "" Then
            BarisTMP = Split(DataTMP(j), delimiterBaris)

            ListLogServer.AddItem TimeValue(Now) & " : Pesan Dari Client : " & Index & " : " & DataTMP(j), 0

            If BarisTMP(0) = "getDataKelas" Then

                Set rs = openRecordset("select * from kelas where deleted = 0")

                ReDim sendTMP(rs.RecordCount + 1)
                sendTMP(0) = "dataKelas"

                i = 1
                If Not rs.EOF Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        sendTMP(i) = rs.Fields("id") & delimiterKolom & rs.Fields("nama")
                        i = i + 1
                        rs.MoveNext
                    Loop
                Else
                    ReDim DataTMP(1)
                    sendTMP(1) = "EOF"
                End If

                dateSend = Join(sendTMP, delimiterBaris)

                wskServer(Index).SendData dateSend

                ListLogServer.AddItem TimeValue(Now) & " : Client " & Index & " meminta data kelas", 0

                Call closeRecordset(rs)

            ElseIf BarisTMP(0) = "getJadwal" Then

'               query = "select jadwal.*, rekapjadwal.id  as rekapjadwalid from jadwal " & _
                        " inner join guru on guru.id = jadwal.idguru " & _
                        " inner join pelajaran on pelajaran.id = jadwal.idpelajaran " & _
                        " inner join rekapjadwal on jadwal.id = rekapjadwal.idjadwal " & _
                        " where idkelas = " & BarisTMP(1) & " " & _
                        " and semester = '" & Me.semester & "' and tahun = '" & Me.tahun & "' " & _
                        " and CAST(FLOOR(CAST(tanggal AS float)) AS datetime) = CAST(FLOOR(CAST(getdate() AS float)) AS datetime)  " & _
                        " and jadwal.waktumulai - CAST(FLOOR(CAST(jadwal.waktumulai AS float)) AS datetime) < '" & CStr(Format$(Now, "hh:mm:ss")) & "' and jadwal.waktuselesai - CAST(FLOOR(CAST(jadwal.waktuselesai AS float)) AS datetime) > '" & CStr(Format$(Now, "hh:mm:ss")) & "' " & " " & _
                        " and guru.deleted = 0 and pelajaran.deleted = 0 and jadwal.deleted = 0"
                        
                query = "select jadwal.*, rekapjadwal.id  as rekapjadwalid from jadwal " & _
                        " inner join guru on guru.id = jadwal.idguru " & _
                        " inner join pelajaran on pelajaran.id = jadwal.idpelajaran " & _
                        " left join rekapjadwal on jadwal.id = rekapjadwal.idjadwal " & _
                        " where idkelas = " & BarisTMP(1) & " " & _
                        " and semester = '" & Me.semester & "' and tahun = '" & Me.tahun & "' " & _
                        " and jadwal.waktumulai < '" & (CDate(CStr(Format$(Now, "hh:mm"))) * 60 * 24) & "' and jadwal.waktuselesai > '" & (CDate(CStr(Format$(Now, "hh:mm"))) * 60 * 24) & "' " & " " & _
                        " and guru.deleted = 0 and pelajaran.deleted = 0 and jadwal.deleted = 0"


                Set rs = openRecordset(query)
                
                If (IsNull(rs.Fields("rekapjadwalid"))) Then
                    Dim rsRekapJadwal As ADODB.Recordset
                    
                    Set rsRekapJadwal = openRecordset("select * from rekapjadwal")
                    
                    rsRekapJadwal.AddNew
                    
                    rsRekapJadwal!idjadwal = rs.Fields("id")
                    rsRekapJadwal!waktumulai = rs.Fields("waktumulai")
                    rsRekapJadwal!waktuselesai = rs.Fields("waktuselesai")
                    rsRekapJadwal!keterangan = "Belum hadir"
                    rsRekapJadwal!deleted = 0
                    
                    rsRekapJadwal.Update
                    
                    Set rs = openRecordset(query)
                End If

                ReDim sendTMP(2)
                sendTMP(0) = "dataJadwal"

                If Not rs.EOF Then
                    rs.MoveFirst
                    sendTMP(1) = rs.Fields("id") & delimiterKolom & rs.Fields("semester") & delimiterKolom & rs.Fields("tahun") & delimiterKolom & rs.Fields("waktumulai") & delimiterKolom & rs.Fields("waktuselesai") & delimiterKolom & rs.Fields("hari") & delimiterKolom & rs.Fields("idkelas") & delimiterKolom & rs.Fields("idpelajaran") & delimiterKolom & rs.Fields("idguru")
                    sendTMP(2) = rs.Fields("rekapjadwalid")
                Else
                    ReDim DataTMP(1)
                    sendTMP(1) = "EOF"
                End If

                dateSend = Join(sendTMP, delimiterBaris)

                wskServer(Index).SendData dateSend

                ListLogServer.AddItem TimeValue(Now) & " : Client " & Index & " meminta data jadwal : " & (CDate(CStr(Format$(Now, "hh:mm"))) * 60 * 24) & " : " & Me.semester & " : " & Me.tahun & " : " & BarisTMP(1) & " : " & CStr(IsNull(rs.Fields("rekapjadwalid"))), 0

                Call closeRecordset(rs)

            ElseIf BarisTMP(0) = "getDataPelajaranAndGuru" Then
                query = "select * from pelajaran where id = " & BarisTMP(1) & " and deleted = 0"

                Set rs = openRecordset(query)

                ReDim sendTMP(2)
                sendTMP(0) = "dataPelajaranDanGuru"


                If Not rs.EOF Then
                    rs.MoveFirst
                    sendTMP(1) = rs.Fields("nama")
                Else
                    sendTMP(UBound(DataTMP)) = "EOF"
                End If

                Call closeRecordset(rs)

                query = "select * from guru where id = " & BarisTMP(2) & " and deleted = 0"

                Set rs = openRecordset(query)

                ReDim Preserve sendTMP(UBound(sendTMP) + 1)

                If Not rs.EOF Then
                    rs.MoveFirst
                    sendTMP(2) = rs.Fields("nama")
                Else
                    sendTMP(UBound(DataTMP)) = "EOF"
                End If

                dateSend = Join(sendTMP, delimiterBaris)

                wskServer(Index).SendData dateSend

                Call closeRecordset(rs)

                ListLogServer.AddItem TimeValue(Now) & " : Client " & Index & " Meminta data pelajaran dan Guru", 0
            ElseIf BarisTMP(0) = "logoutGuru" Then
                Set rs = openRecordset("select * from rekapjadwal where id = '" & BarisTMP(1) & "'")

                If Not rs.EOF Then
                    rs!waktuselesai = (CDate(CStr(Format$(Now, "hh:mm"))) * 60 * 24)
                    rs!keterangan = "Hadir"
                    rs.Update
                End If

                Call closeRecordset(rs)

                ListLogServer.AddItem TimeValue(Now) & " : Guru Log out diclient " & Index, 0
            ElseIf BarisTMP(0) = "loginGuru" Then
                ListLogServer.AddItem TimeValue(Now) & " : Client " & Index & " melakukan login dengan id guru : " & BarisTMP(1) & "' dan nip = '" & Trim$(BarisTMP(2)) & "' dan id rekap jadwal = '" & BarisTMP(3), 0
            
                Set rsDB = openRecordset("select guru.* from guru " & _
                                         " inner join jadwal on jadwal.idguru = guru.id " & _
                                         " inner join rekapjadwal on jadwal.id = rekapjadwal.idjadwal " & _
                                         " where guru.id = '" & BarisTMP(1) & "' and nip = '" & Trim$(BarisTMP(2)) & "' and rekapjadwal.id = '" & BarisTMP(3) & "' " & " and guru.deleted = 0")

                ReDim Preserve sendTMP(2)
                sendTMP(0) = "LoginCheck"

                If rsDB.EOF Then
                    sendTMP(1) = "EOF"
                Else
                    rsDB.MoveFirst
                    sendTMP(1) = rsDB.Fields(0) & delimiterKolom _
                                 & rsDB.Fields(1) & delimiterKolom _
                                 & rsDB.Fields(2) & delimiterKolom _
                                 & rsDB.Fields(3) & delimiterKolom _
                                 & rsDB.Fields(4) & delimiterKolom _
                                 & rsDB.Fields(5) & delimiterKolom _
                                 & rsDB.Fields(6) & delimiterKolom _
                                 & rsDB.Fields(7) & delimiterKolom _
                                 & rsDB.Fields(8) & delimiterKolom _
                                 & rsDB.Fields(9)

                    log = rsDB.Fields("nama")

                    Dim rsSimpan As ADODB.Recordset

                    Set rsSimpan = openRecordset("Select rekapjadwal.* from rekapjadwal " & _
                                                 " inner join jadwal on jadwal.id = rekapjadwal.idjadwal " & _
                                                 " where rekapjadwal.id = '" & BarisTMP(3) & "'  ")


                    If Not rsSimpan.EOF Then
                        Dim rsSimpanRJ As ADODB.Recordset

                        Set rsSimpanRJ = openRecordset("select * from rekapjadwal where id = '" & BarisTMP(3) & "'")

                        If Not rsSimpanRJ.EOF Then
                            rsSimpanRJ!waktumulai = Now
                            rsSimpanRJ!keterangan = "Hadir"
                            rsSimpanRJ.Update
                        End If

                        Call closeRecordset(rsSimpanRJ)

                        sendTMP(2) = rsSimpan.Fields(0) & delimiterKolom _
                                     & rsSimpan.Fields(1) & delimiterKolom _
                                     & rsSimpan.Fields(2) & delimiterKolom _
                                     & rsSimpan.Fields(3) & delimiterKolom _
                                     & rsSimpan.Fields(4) & delimiterKolom _
                                     & rsSimpan.Fields(5)
                    End If

                    Call closeRecordset(rsSimpan)
                End If

                dateSend = Join(sendTMP, delimiterBaris)

                wskServer(Index).SendData dateSend

                Call closeRecordset(rsDB)

                ListLogServer.AddItem TimeValue(Now) & " : Guru Login : " & log, 0
            End If
        End If
    Next j
End Sub

Sub serverError()
    If wskServer(Index).State <> 1 Then wskServer(Index).Close

    lblStatusServer.Caption = "Server Error"

    ListLogServer.AddItem TimeValue(Now) & " : Server Error", 0
End Sub

Sub ServerSendComplete()
    ListLogServer.AddItem TimeValue(Now) & " : Data Telah terkirim", 0
End Sub

Sub SendProgress()
    ListLogServer.AddItem TimeValue(Now) & " : Sedang Mnegirim data", 0
End Sub

Sub ServerShutDown()
    Dim i As Long

    If wskServer(0).State <> sckClosed Then wskServer(0).Close

    For i = 1 To wskServer.UBound
        If wskServer(i).State <> sckClosed Then wskServer(i).Close
        Call Unload(wskServer(i))
    Next i
End Sub

Private Sub TimerSessionManager_Timer()
    If wskServer(0).State = sckListening Then
        Dim rsJadwal As ADODB.Recordset
        Dim rsRekapJadwal As ADODB.Recordset

        query = "select * from jadwal WHERE semester = '" & Me.semester & "' and tahun = '" & Me.tahun & "' and waktumulai - CAST(FLOOR(CAST(waktumulai AS float)) AS datetime) < '" & CStr(Format$(Now, "hh:mm:ss")) & "' and waktuselesai - CAST(FLOOR(CAST(waktuselesai AS float)) AS datetime) > '" & CStr(Format$(Now, "hh:mm:ss")) & "' " & " and deleted = 0"
        Set rsJadwal = openRecordset(query)

        If Not rsJadwal.EOF Then
            rsJadwal.MoveFirst
            Do While Not rsJadwal.EOF
                If rsJadwal.Fields("hari") = UrutanDariHari(Format(Now, "dddd")) Then
                    Set rsRekapJadwal = openRecordset("Select * from rekapjadwal" & _
                                                      " where idjadwal = '" & rsJadwal.Fields(0) & "' " & _
                                                      " and CAST(FLOOR(CAST(tanggal AS float)) AS datetime) = CAST(FLOOR(CAST(getdate() AS float)) AS datetime)")
                    If rsRekapJadwal.EOF Then
                        rsRekapJadwal.AddNew
                        rsRekapJadwal!idjadwal = rsJadwal.Fields(0)
                        rsRekapJadwal!deleted = 0
                        rsRekapJadwal.Update

                        Me.ListLogServer.AddItem TimeValue(Now) & " : Pelajaran dimulai", 0
                    Else
                        Dim rs As ADODB.Recordset

                        If IsNull(rsRekapJadwal.Fields(3).value) = True Then
                            Set rs = openRecordset("select * from rekapjadwal where id = '" & rsRekapJadwal.Fields(0) & "'")
                            If Not rs.EOF Then
                                rs!keterangan = "Tidak Hadir"
                                rs.Update
                            End If

                            Call closeRecordset(rs)
                        Else
                            Set rs = openRecordset("select * from rekapjadwal where id = '" & rsRekapJadwal.Fields(0) & "'")

                            If Not rs.EOF Then
                                rs!waktuselesai = Now
                                rs!keterangan = "Hadir"
                                rs.Update
                            End If

                            Call closeRecordset(rs)
                        End If

                        rsRekapJadwal.Update

                    End If
                    Call closeRecordset(rsRekapJadwal)
                End If
                rsJadwal.MoveNext
            Loop
            Call closeRecordset(rsJadwal)
            Call closeRecordset(rsRekapJadwal)
        End If
    End If
End Sub

Private Sub TimerStatus_Timer()
    Dim rs As ADODB.Recordset

    Set rs = openRecordset("select * from semester where nama = '" & Trim$(cmbSemester.Text) & "'  and deleted  = 0")

    If Not rs.EOF Then
        Me.semester = rs.Fields(0)
    End If

    Call closeRecordset(rs)

    Dim pesan As String
    Dim ip As String
    Dim port As String

    Select Case wskServer(0).State
        Case sckClosing
            pesan = "Tidak terhubung dengan server"
            MsgBox pesan

            ip = ""
            port = ""
            wskServer(0).Close
        Case sckOpen
            pesan = "Saluran Terbuka"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckListening
            pesan = "Saluran Menunggu"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckConnectionPending
            pesan = "Saluran terganggu"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckResolvingHost
            pesan = "Saluran menghilang"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckHostResolved
            pesan = "Saluran server tidak ditemukan"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckConnecting
            pesan = "Saluran menhubungkan"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckConnected
            pesan = "Saluran Tehubung"

            ip = wskServer(0).LocalIP
            port = wskServer(0).LocalPort
        Case sckClosing
            pesan = "Saluran Menutup"

            ip = ""
            port = ""
        Case sckError
            pesan = "Saluran terganggu"

            ip = ""
            port = ""

            wskServer(0).Close
    End Select

    lblStatusServer.Caption = pesan

    Me.lblIPServer.Caption = ip
    Me.lblPortServer.Caption = port

    StatusBar.Panels.Item(1).Text = pesan
    StatusBar.Panels.Item(2).Text = semester
    StatusBar.Panels.Item(3).Text = tahun
    StatusBar.Panels.Item(4).Text = UrutanDariHari(Format(Now, "dddd")) & " - " & Format(Now, "dddd")
End Sub

Private Sub wskServer_Close(Index As Integer)
    Call ServerClose(Index)
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    ServerConnectionRequest Index, requestID
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    ServerDataArrived Index
End Sub


