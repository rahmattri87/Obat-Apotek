VERSION 5.00
Begin VB.Form frm_ctk_Kwitansi_pembayaran 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pembayaran"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ctk_Kwitansi_pembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Text1.Text = frm_pembayaran.txtno_bayar.Text
Text2.Text = frm_pembayaran.txtubay.Text
Text3.Text = frm_pembayaran.txtukem.Text
Dim No As Integer
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from Pembayaran where no_bayar = '" & Text1 & "'"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then
   MsgBox "Data tidak ada?!"
Else
grs1 = String(120, "-")
grs2 = String(120, "=")
frm_ctk_Kwitansi_pembayaran.Font = "Courier New"
frm_ctk_Kwitansi_pembayaran.FontSize = 12
frm_ctk_Kwitansi_pembayaran.FontBold = False
Printer.Print Tab(0)
Printer.Print Tab(0)
Printer.Print Tab(5); "APOTIK ASTITI"
Me.FontSize = 12
Printer.Print Tab(5); "Jl. Siswa No.2 Tangerang"
Printer.Print Tab(5); "Telp. (021)552495"
Printer.Print Tab(5); "----------------------------------------------------------"
frm_ctk_Kwitansi_pembayaran.FontSize = 10
Printer.Print Tab(6); "NO.Bayar         :"; Data1.Recordset!no_bayar; "     "; "Kode Dokter    :"; Data1.Recordset!kode_dokter
Printer.Print Tab(6); "Tanggal Bayar    :"; Data1.Recordset!tgl_byr; "    "; "Nama Dokter    :"; Data1.Recordset!nama_dokter
Printer.Print Tab(6); "No.Resep         :"; Data1.Recordset!No_resep; "     "; "Kode Pasien    :"; Data1.Recordset!kode_pasien
Printer.Print Tab(6); "Tanggal Resep    :"; Data1.Recordset!tgl_resep; "    "; "Nama Pasien    :"; Data1.Recordset!nama_Pasien
Printer.Print Tab(0)
Printer.FontSize = 10
Printer.Print Tab(6); "========================================================================"
Printer.Print Tab(6); "No."; "  Kode Obat"; "    Nama Obat"; "    Jumlah Obat"; "     Satuan"; "          Bayar"
Printer.Print Tab(6); "========================================================================"
If Not Data1.Recordset.RecordCount = 0 Then
       Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
    
No = No + 1
hit = hit + Data1.Recordset!jumlah
Printer.Print Tab(5); No; ".   "; Data1.Recordset!kode_obat; vbTab; Data1.Recordset!nama_obat; vbTab; vbTab; Data1.Recordset!jml_obat; vbTab; "Rp. "; Format(Data1.Recordset!harga_satuan, "###,###.-"); vbTab; "Rp. "; Format(Data1.Recordset!jumlah, "###,###.-")
Data1.Recordset.MoveNext
Loop
End If
Printer.Print Tab(6); "========================================================================"
Printer.Print Spc(45); "Total Bayar    : "; Spc(2); "Rp. "; Format(hit, "###,###.-")
Printer.Print Spc(45); "Uang Bayar     : "; Spc(2); "Rp. "; Format(Text2, "###,###.-")
Printer.Print Spc(45); "Uang Kembali   : "; Spc(2); "Rp. "; Format(Text3, "###,###.-")
Printer.EndDoc
End If
End Sub

Private Sub Form_Activate()
Text1.Text = frm_pembayaran.txtno_bayar.Text
Text2.Text = frm_pembayaran.txtubay.Text
Text3.Text = frm_pembayaran.txtukem.Text
Dim No As Integer
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from Pembayaran where no_bayar = '" & Text1 & "'"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then
   MsgBox "Data tidak ada?!"
Else
grs1 = String(120, "-")
grs2 = String(120, "=")
frm_ctk_Kwitansi_pembayaran.Font = "Courier New"
frm_ctk_Kwitansi_pembayaran.FontSize = 12
frm_ctk_Kwitansi_pembayaran.FontBold = False
Print Tab(0)
Print Tab(0)
Print Tab(5); "APOTIK ASTITI"
Me.FontSize = 12
Print Tab(5); "Jl. Siswa No.2 Tangerang"
Print Tab(5); "Telp. (021)552495"
Print Tab(5); "----------------------------------------------------------"
frm_ctk_Kwitansi_pembayaran.FontSize = 10
Print Tab(6); "NO.Bayar         :"; Data1.Recordset!no_bayar; "     "; "Kode Dokter    :"; Data1.Recordset!kode_dokter
Print Tab(6); "Tanggal Bayar    :"; Data1.Recordset!tgl_byr; "    "; "Nama Dokter    :"; Data1.Recordset!nama_dokter
Print Tab(6); "No.Resep         :"; Data1.Recordset!No_resep; "     "; "Kode Pasien    :"; Data1.Recordset!kode_pasien
Print Tab(6); "Tanggal Resep    :"; Data1.Recordset!tgl_resep; "    "; "Nama Pasien    :"; Data1.Recordset!nama_Pasien
Print Tab(0)
Printer.FontSize = 10
Print Tab(6); "========================================================================"
Print Tab(6); "No."; "  Kode Obat"; "    Nama Obat"; "    Jumlah Obat"; "     Satuan"; "          Bayar"
Print Tab(6); "========================================================================"
If Not Data1.Recordset.RecordCount = 0 Then
       Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
    
No = No + 1
hit = hit + Data1.Recordset!jumlah
Print Tab(5); No; ".   "; Data1.Recordset!kode_obat; vbTab; Data1.Recordset!nama_obat; vbTab; vbTab; Data1.Recordset!jml_obat; vbTab; "Rp. "; Format(Data1.Recordset!harga_satuan, "###,###.-"); vbTab; "Rp. "; Format(Data1.Recordset!jumlah, "###,###.-")
Data1.Recordset.MoveNext
Loop
End If
Print Tab(6); "========================================================================"
Print Spc(45); "Total Bayar    : "; Spc(2); "Rp. "; Format(hit, "###,###.-")
Print Spc(45); "Uang Bayar     : "; Spc(2); "Rp. "; Format(Text2, "###,###.-")
Print Spc(45); "Uang Kembali   : "; Spc(2); "Rp. "; Format(Text3, "###,###.-")
End If
End Sub

Private Sub Form_Click()
Unload Me
End Sub
