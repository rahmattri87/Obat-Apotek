VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form lap_pembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN DATA PEMBAYARAN"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Lihat Laporan Pembayaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "DD-MM-YYYY"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "DD-MM-YYYY"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb;Persist Security Info=False"
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
End
Attribute VB_Name = "lap_pembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If con.State = 0 Then
    con.Open
End If
con.Execute "delete from Lap_Pembayaran"
SQL = "insert into Lap_Pembayaran select no_resep,tgl_resep,kode_dokter,nama_dokter,kode_pasien,nama_pasien,kode_obat,nama_obat, jml_obat,harga_satuan,jumlah,no_bayar,tgl_byr from pembayaran where tgl_byr between #" & Text1.Text & "# and #" & Text2.Text & "#"
''#" & Text1.Text & "# and #" & Text2.Text & "#"
Debug.Print SQL
con.Execute SQL
con.Close
If DataEnvironment1.rsLap_Pembayaran_Bulanan_Grouping.State = 0 Then
    DataEnvironment1.rsLap_Pembayaran_Bulanan_Grouping.Open
End If
DataEnvironment1.rsLap_Pembayaran_Bulanan_Grouping.Requery
Lap_Pembayaran_Bulanan.Refresh
Lap_Pembayaran_Bulanan.Show
DataEnvironment1.rsLap_Pembayaran_Bulanan_Grouping.Close
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb;Persist Security Info=False"
    con.Open
End Sub
'

