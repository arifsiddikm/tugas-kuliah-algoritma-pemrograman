VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16230
   LinkTopic       =   "Form14"
   ScaleHeight     =   10980
   ScaleWidth      =   16230
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   18865
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tambah Data Penjualan"
      TabPicture(0)   =   "Program UAS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "btnHitung"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "formKodeBus"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "formKodeKelas"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "formJumlahTiket"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "formNamaBus"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "formJurusan"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "formKelas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "formTotal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "formTarifDasar"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "formDiskon"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "formNamaPenumpang"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "btnTutup"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "btnHapus"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "formNoPenumpang"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "formTarifTambahan"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "formBayar"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "btnTambah"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Tampilkan Data Penjualan"
      TabPicture(1)   =   "Program UAS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "Adodc1"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton btnTambah 
         BackColor       =   &H8000000D&
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         MaskColor       =   &H00FFFF00&
         TabIndex        =   32
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox formBayar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   30
         Top             =   9360
         Width           =   10695
      End
      Begin VB.TextBox formTarifTambahan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   28
         Top             =   6840
         Width           =   10695
      End
      Begin VB.TextBox formNoPenumpang 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1920
         TabIndex        =   25
         Text            =   "P000"
         Top             =   1800
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   -74760
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   9015
         Left            =   -74760
         TabIndex        =   24
         Top             =   1440
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   15901
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton btnHapus 
         BackColor       =   &H8000000D&
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         MaskColor       =   &H00FFFF00&
         TabIndex        =   13
         Top             =   3480
         Width           =   3495
      End
      Begin VB.CommandButton btnTutup 
         BackColor       =   &H8000000D&
         Caption         =   "TUTUP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         MaskColor       =   &H00FFFF00&
         TabIndex        =   12
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox formNamaPenumpang 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   6240
         TabIndex        =   11
         Text            =   "Arif"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox formDiskon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   10
         Top             =   8520
         Width           =   10695
      End
      Begin VB.TextBox formTarifDasar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   9
         Top             =   6000
         Width           =   10695
      End
      Begin VB.TextBox formTotal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   8
         Top             =   7680
         Width           =   10695
      End
      Begin VB.TextBox formKelas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   9360
         TabIndex        =   7
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox formJurusan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   6
         Top             =   5160
         Width           =   10695
      End
      Begin VB.TextBox formNamaBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2760
         TabIndex        =   5
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox formJumlahTiket 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   11880
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox formKodeKelas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         ItemData        =   "Program UAS.frx":0038
         Left            =   9360
         List            =   "Program UAS.frx":0048
         TabIndex        =   3
         Text            =   "BS"
         Top             =   2640
         Width           =   4095
      End
      Begin VB.ComboBox formKodeBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         ItemData        =   "Program UAS.frx":005C
         Left            =   2760
         List            =   "Program UAS.frx":006F
         TabIndex        =   2
         Text            =   "GO"
         Top             =   2640
         Width           =   4095
      End
      Begin VB.CommandButton btnHitung 
         BackColor       =   &H8000000D&
         Caption         =   "HITUNG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         MaskColor       =   &H00FFFF00&
         TabIndex        =   1
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   " BAYAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   9360
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "TARIF TAMBAHAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   " PENUMPANG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3720
         TabIndex        =   27
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " NOMOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   " DISKON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   8520
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Caption         =   "TARIF DASAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   " TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   7680
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "KELAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6840
         TabIndex        =   20
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   " JURUSAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   " NAMA BUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KODE KELAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6840
         TabIndex        =   17
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "JUMLAH TIKET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   9480
         TabIndex        =   16
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "KODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "PROGRAM PENJUALAN TIKET BUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   13215
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KoneksiDB As New ADODB.Connection
Dim Bayar As ADODB.Recordset

Sub BukaDB()
  Set KoneksiDB = New ADODB.Connection
  Set Bayar = New ADODB.Recordset
  KoneksiDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\UAS_ArifSiddikMuharam.mdb;"
End Sub

Private Sub btnHapus_Click()
    Call BukaDB
    Dim tambahdataa
    Dim cekdata
    Dim sqlcekdata
    Dim hapusdata
    cekdata = "SELECT * FROM Bayar WHERE NoPenumpang = '" & formNoPenumpang & "'"
    Set sqlcekdata = KoneksiDB.Execute(cekdata)
    If Not sqlcekdata.EOF Then
        If MsgBox("Data dengan No Penumpang " + formNoPenumpang + " sudah tersedia, Apakah Anda ingin menghapus data ini dari Database?", 36, "Informasi") = vbYes Then
          hapusdata = "DELETE FROM Bayar WHERE NoPenumpang = '" & formNoPenumpang & "'"
          KoneksiDB.Execute (hapusdata)
          MsgBox "Data Penjualan Tiket Bus Berhasil Dihapus ^_^"
          Adodc1.Refresh
        End If
    End If
End Sub

Private Sub btnHitung_Click()
  Dim formTarifDasar1
  Dim formTotal1
  Dim formTarifTambahan1
  Dim formDiskon1
  Dim formBayar1
  If formJumlahTiket = "" Then
    formNoPenumpang = "P000"
    formNamaPenumpang = "Arif"
    formKodeBus = "GO"
    formKodeKelas = "BS"
    formNamaBus = ""
    formKelas = ""
    formJurusan = ""
    formTarifDasar = ""
    formTarifTambahan = ""
    formTotal = ""
    formDiskon = ""
    formBayar = ""
  Else
    If formKodeBus = "GO" Then
        formTarifDasar1 = 80000
        formTarifDasar = "Rp. 80.000"
        formNamaBus = "P.O. Sahabat"
        formJurusan = "Jakarta - Cirebon"
    ElseIf formKodeBus = "KU" Then
        formTarifDasar1 = 300000
        formTarifDasar = "Rp. 300.000"
        formNamaBus = "P.O. Kramat Jati"
        formJurusan = "Jakarta - Denpasar"
    ElseIf formKodeBus = "SH" Then
        formTarifDasar1 = 100000
        formTarifDasar = "Rp. 100.000"
        formNamaBus = "P.O. Hiba Utama"
        formJurusan = "Jakarta - Kuningan"
    ElseIf formKodeBus = "SJ" Then
        formTarifDasar1 = 250000
        formTarifDasar = "Rp. 250.000"
        formNamaBus = "P.O. Tali Jaya"
        formJurusan = "Jakarta - Surabaya"
    ElseIf formKodeBus = "SN" Then
        formTarifDasar1 = 200000
        formTarifDasar = "Rp. 200.000"
        formNamaBus = "P.O. Pahala Kencana"
        formJurusan = "Jakarta - Semarang"
    Else
        formTarifDasar1 = 0
        formTarifDasar = ""
        formNamaBus = ""
        formJurusan = ""
    End If
    formTotal1 = Val(formTarifDasar1) * Val(formJumlahTiket)
    formTotal = "Rp. " + FormatNumber(formTotal1, 0)
    If formKodeKelas = "BS" Then
        formKelas = "Bisnis"
        formTarifTambahan1 = 15 / 100 * Val(formTotal1)
        formTarifTambahan = "Rp. " + FormatNumber(formTarifTambahan1, 0)
        formDiskon1 = 5 / 100
        formDiskon = "5%"
    ElseIf formKodeKelas = "EK" Then
        formKelas = "Ekonomi"
        formTarifTambahan1 = 15 / 100 * Val(formTotal1)
        formTarifTambahan = "Rp. " + FormatNumber(formTarifTambahan1, 0)
        formDiskon1 = 5 / 100
        formDiskon = "5%"
    ElseIf formKodeKelas = "ES" Then
        formKelas = "Eksekutif"
        formTarifTambahan1 = 20 / 100 * Val(formTotal1)
        formTarifTambahan = "Rp. " + FormatNumber(formTarifTambahan1, 0)
        formDiskon1 = 10 / 100
        formDiskon = "10%"
    ElseIf formKodeKelas = "SE" Then
        formKelas = "Super Eksekutif"
        formTarifTambahan1 = 25 / 100 * Val(formTotal1)
        formTarifTambahan = "Rp. " + FormatNumber(formTarifTambahan1, 0)
        formDiskon1 = 10 / 100
        formDiskon = "10%"
    Else
        formKelas = ""
        formTarifTambahan1 = ""
        formTarifTambahan = ""
        formDiskon1 = 0
        formDiskon = ""
    End If
    formBayar1 = Val(formTotal1) + Val(formTarifTambahan1) - Val(formDiskon1)
    formBayar = "Rp. " + FormatNumber(formBayar1, 0)
  End If
End Sub

Private Sub btnTambah_Click()
    Call BukaDB
    Dim tambahdataa
    Dim cekdata
    Dim sqlcekdata
    cekdata = "SELECT * FROM Bayar WHERE NoPenumpang = '" & formNoPenumpang & "'"
    Set sqlcekdata = KoneksiDB.Execute(cekdata)
    If Not sqlcekdata.EOF Then
        MsgBox "Data dengan No Penumpang " + formNoPenumpang + " sudah tersedia, harap masukkan No Penumpang yang baru ^_^"
    Else
        tambahdataa = "Insert into Bayar values('" & formNoPenumpang & "','" & formNamaPenumpang & "','" & formKodeBus & "','" & formKodeKelas & "','" & formJumlahTiket & "')"
        KoneksiDB.Execute (tambahdataa)
        MsgBox "Data Penjualan Tiket Bus Berhasil Ditambahkan ^_^"
        Adodc1.Refresh
    End If
End Sub

Private Sub Form_Load()
  Call BukaDB
  Adodc1.ConnectionString = KoneksiDB
  Adodc1.RecordSource = "Bayar"
  Adodc1.Refresh
  Set DataGrid1.DataSource = Adodc1
End Sub
