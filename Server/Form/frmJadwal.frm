VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmJadwal 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   14737
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Mata Pelajaran"
      TabPicture(0)   =   "frmJadwal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Kelas"
      TabPicture(1)   =   "frmJadwal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "lstKelas"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "cmdTambahKelas"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Jam Pelajaran"
      TabPicture(2)   =   "frmJadwal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdTambahJamPelajaran"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "lstJamPelajaran"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame4 
         Caption         =   "Pencarian Data Jam Pelajaran"
         Height          =   855
         Left            =   -74820
         TabIndex        =   7
         Top             =   420
         Width           =   9795
         Begin VB.CommandButton Command6 
            Caption         =   "Cari"
            Height          =   375
            Left            =   8280
            TabIndex        =   9
            Top             =   300
            Width           =   1395
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1860
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   300
            Width           =   6255
         End
         Begin VB.Label Label12 
            Caption         =   "Pencarian"
            Height          =   375
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdTambahJamPelajaran 
         Caption         =   "Tambah Guru"
         Height          =   375
         Left            =   -66480
         TabIndex        =   6
         Top             =   7860
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pencarian Data Kelas"
         Height          =   855
         Left            =   -74820
         TabIndex        =   2
         Top             =   420
         Width           =   9795
         Begin VB.CommandButton Command4 
            Caption         =   "Cari"
            Height          =   375
            Left            =   8280
            TabIndex        =   4
            Top             =   300
            Width           =   1395
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1860
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   300
            Width           =   6255
         End
         Begin VB.Label Label9 
            Caption         =   "Pencarian"
            Height          =   375
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdTambahKelas 
         Caption         =   "Tambah Kelas"
         Height          =   375
         Left            =   -66480
         TabIndex        =   1
         Top             =   7860
         Width           =   1395
      End
      Begin MSComctlLib.ListView lstKelas 
         Height          =   6360
         Left            =   -74820
         TabIndex        =   11
         Top             =   1305
         Width           =   9780
         _ExtentX        =   17251
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
      Begin MSComctlLib.ListView lstJamPelajaran 
         Height          =   6360
         Left            =   -74820
         TabIndex        =   12
         Top             =   1305
         Width           =   9780
         _ExtentX        =   17251
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
      Begin VB.Label Label11 
         Caption         =   "Jumlah Guru : "
         Height          =   255
         Left            =   -74820
         TabIndex        =   14
         Top             =   7860
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "lblJumlahGuru"
         Height          =   255
         Left            =   -73740
         TabIndex        =   13
         Top             =   7860
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmJadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
