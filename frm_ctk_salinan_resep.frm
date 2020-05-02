VERSION 5.00
Begin VB.Form frm_ctk_salinan_resep 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Resep"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tutup"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Cetak"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtno_resep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaxLength       =   8
      TabIndex        =   0
      Top             =   7080
      Width           =   2175
   End
End
Attribute VB_Name = "frm_ctk_salinan_resep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim No As Integer
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from resep where no_resep = '" & txtno_resep & "'"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then
   MsgBox "Data tidak ada?!"
Else
grs1 = String(120, "-")
grs2 = String(120, "=")
frm_ctk_salinan_resep.Font = "Courier New"
Me.FontSize = 12
Me.FontBold = False
Printer.Print Tab(0)
Printer.Print Tab(0)
Printer.Print Tab(5); "APOTIK ASTITI"
'Me.FontSize = 12
Printer.Print Tab(5); "Jl. Siswa No.2 Tangerang"
Printer.Print Tab(5); "Telp. (021)552495"
Printer.Print Tab(5); "----------------------------------------------------------"
Me.FontSize = 10
Printer.Print Tab(6); "NO.Resep         : "; Data1.Recordset!No_resep
Printer.Print Tab(6); "Nama dokter     : "; Data1.Recordset!nama_dokter
Printer.Print Tab(6); "Tgl. Buat           : "; Data1.Recordset!tgl
Printer.Print Tab(6); "Untuk                : "; Data1.Recordset!nama_pasien
Printer.Print Tab(6); "Umur                 : "; Data1.Recordset!Umur
Printer.Print Tab(0)
Me.FontSize = 10
Printer.Print Tab(6); "========================================================================"
Printer.Print Tab(6); "No."; "  Kode Obat"; "    Nama Obat"; "    Jml Obat"; "       Satuan"; "              Bayar"
Printer.Print Tab(6); "========================================================================"
If Not Data1.Recordset.RecordCount = 0 Then
       Data1.Recordset.MoveFirst
    Do While Not Data1.Recordset.EOF
    
No = No + 1
hit = hit + Data1.Recordset!jumlah
Printer.Print Tab(5); No; ".   "; Data1.Recordset!Kode_Obat; vbTab; vbTab; Data1.Recordset!Nama_Obat; vbTab; vbTab; Data1.Recordset!jml_obat; vbTab; "Rp. "; Format(Data1.Recordset!harga_satuan, "###,###.-"); vbTab; "Rp. "; Format(Data1.Recordset!jumlah, "###,###.-")
Data1.Recordset.MoveNext
Loop
End If
Printer.Print Tab(6); "========================================================================"
Printer.Print Spc(49); "Total Bayar : "; Spc(1); "Rp. "; Format(hit, "###,###.-")
Printer.EndDoc
End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
txtno_resep = ""
txtno_resep.Enabled = True
frm_ctk_salinan_resep.Cls
Command1.Enabled = False
End Sub

Private Sub Form_Activate()
Command1.Enabled = False
End Sub

Private Sub txtno_resep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
Command1.Enabled = True
End If

End Sub

Private Sub txtno_resep_LostFocus()
Dim No As Integer
Dim TotalBayar As Long
Dim grs1 As String
Dim grs2 As String
No = 0
TotalBayar = 0

Data1.RecordSource = "select * from resep where no_resep = '" & txtno_resep & "'"
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
Print Tab(5); "APOTIK ASTITI"
Me.FontSize = 12
Print Tab(5); "Jl. Siswa No.2 Tangerang"
Print Tab(5); "Telp. (021)552495"
Print Tab(5); "----------------------------------------------------------"
frm_ctk_salinan_resep.FontSize = 10
Print Tab(6); "NO.Resep         :"; Data1.Recordset!No_resep
Print Tab(6); "Nama dokter      :"; Data1.Recordset!nama_dokter
Print Tab(6); "Tgl. Buat        :"; Data1.Recordset!tgl
Print Tab(6); "Untuk            :"; Data1.Recordset!nama_pasien
Print Tab(6); "Umur             :"; Data1.Recordset!Umur
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
Print Tab(5); No; ".   "; Data1.Recordset!Kode_Obat; vbTab; Data1.Recordset!Nama_Obat; vbTab; vbTab; Data1.Recordset!jml_obat; vbTab; "Rp. "; Format(Data1.Recordset!harga_satuan, "###,###.-"); vbTab; "Rp. "; Format(Data1.Recordset!jumlah, "###,###.-")
Data1.Recordset.MoveNext
Loop
End If
Print Tab(6); "========================================================================"
Print Spc(49); "Total Bayar : "; Spc(1); "Rp. "; Format(hit, "###,###.-")

End If
End Sub
