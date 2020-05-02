VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_obat_bebas 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OBAT BEBAS"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtstockawal 
      Height          =   285
      Left            =   9720
      TabIndex        =   34
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtstockakhir 
      Height          =   285
      Left            =   9720
      TabIndex        =   33
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9720
      TabIndex        =   32
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9720
      TabIndex        =   31
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9720
      TabIndex        =   30
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9720
      TabIndex        =   29
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Obat_Bebas"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Bebas_Sementara"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Obat"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000000&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   5760
      Width           =   9255
      Begin VB.CommandButton cmd_Simpan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Simpan Data"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmd_Keluar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Keluar"
         Height          =   615
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   35
         Top             =   1800
         Width           =   975
      End
      Begin MSDBGrid.DBGrid Grid 
         Bindings        =   "frm_obat_bebas.frx":0000
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frm_obat_bebas.frx":0014
         TabIndex        =   28
         Top             =   2280
         Width           =   7815
      End
      Begin VB.CommandButton cmd_tambah 
         BackColor       =   &H80000009&
         Caption         =   "Tambah &Obat"
         Height          =   615
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtjml_obat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtNama_obat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtharga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtNo_struk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtTanggal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dddd, d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtkd_obat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmd_Hapus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Hapus Obat"
         Height          =   615
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000007&
         Caption         =   "Label18"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4080
         Width           =   7815
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000007&
         Caption         =   "Label16"
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   7815
      End
      Begin VB.Label Label15 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Satuan"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Struk"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal "
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.TextBox txtBayar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtkembalian 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   27
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   26
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kembalian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   25
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "frm_obat_bebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub buatnostruk()
If Data3.Recordset.RecordCount <> 0 Then
        Data3.Recordset.MoveLast
        txtNo_struk = Right(Data3.Recordset!no_struk, 3)
        Nostr = Val(txtNo_struk) + 1
        Select Case Nostr
            Case 0 To 9
                txtNo_struk.Text = "S0000" & Trim(Str(Nostr))
            Case 10 To 99
                txtNo_struk.Text = "S000" & Trim(Str(Nostr))
            Case 100 To 999
                txtNo_struk.Text = "S00" & Trim(Str(Nostr))
            Case 1000 To 9999
                txtNo_struk.Text = "S0" & Trim(Str(Nostr))
            Case 10000 To 99999
                txtNo_struk.Text = "S" & Trim(Str(Nostr))
        End Select
    Else
        txtNo_struk.Text = "S00001"
    End If

End Sub
Sub kosong()
txtkd_obat = ""
txtNama_obat = ""
txtjml_obat = ""
txtharga = ""
txtJumlah = ""
txtTotal = ""
txtBayar = ""
txtkembalian = ""
End Sub

Private Sub cmd_Batal_Click()

End Sub

Private Sub cmd_Hapus_Click()
If Data2.Recordset.RecordCount = 0 Then
   MsgBox "Maaf Tidak ada Yang Harus dibatalkan", vbInformation, "Info"
   Grid.Refresh
   Exit Sub
End If
x = MsgBox("Betul data akan dihapus ?", vbOKCancel + vbInformation, "Menghapus Record")
    If x = vbOK Then
        
         Data1.Recordset.Index = "Kode_obat"
            Data1.Recordset.Seek "=", Text1.Text
            If Not Data1.Recordset.NoMatch Then
                Data1.Recordset.edit
                Data1.Recordset!jml_obat = Val(Data1.Recordset!jml_obat) + Val(Text2.Text)
                txtTotal.Text = Val(txtTotal.Text) - Val(Text3.Text)
                Data1.Recordset.Update
            End If
         Data2.Recordset.delete
         Data2.Recordset.MoveFirst
         Grid.Refresh
         Text1.Text = ""
         Text2.Text = ""
         Text3.Text = ""
    End If
End Sub

Private Sub cmd_Keluar_Click()
Unload Me
End Sub


Private Sub cmd_Simpan_Click()
If Data2.Recordset.RecordCount = 0 Then
   MsgBox "maaf data harus lengkap"
   Exit Sub
Else
  If Data2.Recordset.RecordCount > 0 Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
Data3.Recordset.AddNew
Data3.Recordset!no_struk = txtNo_struk
Data3.Recordset!tgl_struk = txtTanggal.Text
Data3.Recordset!kode_obat = Data2.Recordset!kode_obat
Data3.Recordset!nama_obat = Data2.Recordset!nama_obat
Data3.Recordset!jml_obat = Data2.Recordset!jml_obat
Data3.Recordset!harga_satuan = Data2.Recordset!harga_obat
Data3.Recordset!jumlah = Data2.Recordset!jumlah
Data3.Recordset.Update
            Data2.Recordset.MoveNext
        Loop
    End If
    If Data2.Recordset.RecordCount > 0 Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.delete
            Data2.Recordset.MoveNext
        Loop
    End If
 x = MsgBox("mau cetak struk?", vbInformation + vbYesNo, "Info")
 If x = vbYes Then
 frm_ctk_struk.Show
 Else
Call kosong
cmd_Simpan.Enabled = False
buatnostruk
End If
End If
End Sub

Private Sub cmd_tambah_Click()
If txtkd_obat.Text = "" Then
   MsgBox "isi dulu kode obatnya", vbInformation, "info"
   txtkd_obat.SetFocus
   Exit Sub
End If
   Data2.Recordset.AddNew
   Data2.Recordset!kode_obat = txtkd_obat.Text
   Data2.Recordset!nama_obat = txtNama_obat.Text
   Data2.Recordset!jml_obat = txtjml_obat.Text
   Data2.Recordset!harga_obat = txtharga.Text
   Data2.Recordset!jumlah = txtJumlah.Text
   Data2.Recordset.Update
   
            Data1.Recordset.Index = "Kode_obat"
            Data1.Recordset.Seek "=", txtkd_obat
          
            If Not Data1.Recordset.NoMatch Then
                Data1.Recordset.edit
                Data1.Recordset!jml_obat = Val(Data1.Recordset!jml_obat) - Val(txtjml_obat.Text)
                Data1.Recordset.Update
            End If
            txtTotal = Val(txtTotal.Text) + Val(txtJumlah.Text)
   txtkd_obat.Text = ""
   txtNama_obat.Text = ""
   txtjml_obat.Text = ""
   txtharga.Text = ""
   txtJumlah.Text = ""

End Sub

Private Sub Form_Activate()
txtTanggal = Format(Date, "DD-MM-YYYY")
buatnostruk
cmd_Simpan.Enabled = False
End Sub

Private Sub Grid_Click()
Text1.Text = Grid.Columns(0)
Text2.Text = Grid.Columns(2)
Text3.Text = Grid.Columns(4)
End Sub

Private Sub Grid_DblClick()
Text1.Text = Grid.Columns(0)
Text2.Text = Grid.Columns(2)
Text3.Text = Grid.Columns(4)
End Sub

Private Sub txtBayar_Change()
txtkembalian = Val(txtBayar.Text) - Val(txtTotal.Text)
If txtkembalian.Text >= 0 Then
   cmd_Simpan.Enabled = True
Else
   cmd_Simpan.Enabled = False
End If

End Sub

Private Sub txtjml_obat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If CDate(Text5.Text) <= CDate(txtTanggal.Text) Then
   MsgBox "Maaf Obat Ini Sudah Kadaluarsa", vbInformation, "info"
   txtkd_obat = ""
   txtNama_obat = ""
   txtjml_obat = ""
   txtharga = ""
   txtJumlah = ""
   Text5 = ""
   Exit Sub

Else
      txtJumlah = txtharga * Val(txtjml_obat)
      cmd_tambah.SetFocus
End If
End If
End Sub

Private Sub txtkd_obat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_browse_obat_bebas.Show
frm_browse_obat_bebas.Grid.Refresh
End If
End Sub
