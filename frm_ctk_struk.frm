VERSION 5.00
Begin VB.Form frm_ctk_struk 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Obat_Bebas"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ctk_struk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from obat_bebas where no_struk = '" & Text1.Text & "'"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then
   MsgBox "Data tidak ada?!"
Else
grs1 = String(120, "-")
grs2 = String(120, "=")
frm_ctk_salinan_resep.Font = "Courier New"
frm_ctk_salinan_resep.FontSize = 12
frm_ctk_salinan_resep.FontBold = False
Printer.Print Tab(0)
Printer.Print Tab(0)
Me.FontSize = 12
Printer.Print Tab(5); "APOTIK ASTITI"
Printer.Print Tab(5); "Jl. Siswa No.2 Tangerang"
Printer.Print Tab(5); "Telp. (021)552495"
Printer.Print Tab(5); "------------------------------------------------"
frm_ctk_salinan_resep.FontSize = 10
Printer.Print Tab(5); "NO.Struk          :"; Data1.Recordset!no_struk
Printer.Print Tab(5); "Tanggal Struk  :"; Data1.Recordset!tgl_struk
Printer.Print Tab(0)
Printer.FontSize = 10
If Not Data1.Recordset.RecordCount = 0 Then
       Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
No = No + 1
hit = hit + Data1.Recordset!jumlah
Printer.Print Tab(5); No; ".   "; Data1.Recordset!kode_obat; "  "; Data1.Recordset!nama_obat;
Printer.Print Tab(9); Data1.Recordset!harga_satuan; " X "; Data1.Recordset!jml_obat; vbTab; "    = "; Data1.Recordset!jumlah
Data1.Recordset.MoveNext
Loop
End If
Printer.Print Tab(6); "==========================="
Printer.Print Spc(8); "Total Bayar      : "; Spc(2); "Rp. "; Format(hit, "###,###.-")
Printer.Print Spc(8); "Uang Bayar     : "; Spc(2); "Rp. "; Format(Text2, "###,###.-")
Printer.Print Spc(8); "Uang Kembali : "; Spc(2); "Rp. "; Format(Text3, "###,###.-")
Printer.EndDoc
End If
End Sub

Private Sub Form_Activate()
Text1.Text = frm_obat_bebas.txtNo_struk.Text
Text2.Text = frm_obat_bebas.txtBayar
Text3.Text = frm_obat_bebas.txtkembalian
Dim No As Integer
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from obat_bebas where no_struk = '" & Text1.Text & "'"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then
   MsgBox "Data tidak ada?!"
Else
grs1 = String(120, "-")
grs2 = String(120, "=")
frm_ctk_salinan_resep.Font = "Courier New"
frm_ctk_salinan_resep.FontSize = 12
frm_ctk_salinan_resep.FontBold = False
Print Tab(0)
Print Tab(0)
Me.FontSize = 12
Print Tab(5); "APOTIK ASTITI"
Print Tab(5); "Jl. Siswa No.2 Tangerang"
Print Tab(5); "Telp. (021)552495"
Print Tab(5); "------------------------------------------------"
frm_ctk_salinan_resep.FontSize = 10
Print Tab(5); "NO.Struk          :"; Data1.Recordset!no_struk
Print Tab(5); "Tanggal Struk  :"; Data1.Recordset!tgl_struk
Print Tab(0)
Printer.FontSize = 10
If Not Data1.Recordset.RecordCount = 0 Then
       Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
No = No + 1
hit = hit + Data1.Recordset!jumlah
Print Tab(5); No; ".   "; Data1.Recordset!kode_obat; "  "; Data1.Recordset!nama_obat;
Print Tab(9); Data1.Recordset!harga_satuan; " X "; Data1.Recordset!jml_obat; vbTab; "    = "; Data1.Recordset!jumlah
Data1.Recordset.MoveNext
Loop
End If
Print Tab(6); "==========================="
Print Spc(8); "Total Bayar      : "; Spc(2); "Rp. "; Format(hit, "###,###.-")
Print Spc(8); "Uang Bayar     : "; Spc(2); "Rp. "; Format(Text2, "###,###.-")
Print Spc(8); "Uang Kembali : "; Spc(2); "Rp. "; Format(Text3, "###,###.-")
End If
End Sub

Private Sub Form_Click()
frm_obat_bebas.txtTotal.Text = ""
frm_obat_bebas.txtBayar.Text = ""
frm_obat_bebas.txtkembalian.Text = ""
frm_obat_bebas.cmd_Simpan.Enabled = False
Unload Me
End Sub

