VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_resep 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resep"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9720
      TabIndex        =   42
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9720
      TabIndex        =   41
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9720
      TabIndex        =   40
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9720
      TabIndex        =   39
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtstockakhir 
      Height          =   285
      Left            =   9720
      TabIndex        =   37
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtstockawal 
      Height          =   285
      Left            =   9720
      TabIndex        =   36
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Resep"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Resep_Sementara"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Obat"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Pasien"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Dokter"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000000&
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   6120
      Width           =   9135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Isi Data Resep"
         Height          =   615
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmd_Simpan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Simpan Data"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmd_Keluar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Keluar"
         Height          =   615
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalinanResep 
         BackColor       =   &H80000009&
         Caption         =   "Cetak Salinan Resep"
         Height          =   615
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Rp""#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   2
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
      Height          =   330
      Left            =   6000
      TabIndex        =   30
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   43
         Top             =   2760
         Width           =   975
      End
      Begin MSDBGrid.DBGrid Grid 
         Bindings        =   "frm_resep.frx":0000
         Height          =   1935
         Left            =   120
         OleObjectBlob   =   "frm_resep.frx":0014
         TabIndex        =   29
         Top             =   3240
         Width           =   7815
      End
      Begin VB.TextBox txtumur_Pasien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmd_tambah 
         BackColor       =   &H80000009&
         Caption         =   "Tambah &Obat"
         Height          =   615
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6480
         TabIndex        =   12
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtjml_obat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtNama_obat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtharga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtnama_Pasien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtnama_dokter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtNo_resep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   960
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
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtkd_obat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtkd_dokter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtkd_pasien 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Hapus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Hapus Obat"
         Height          =   615
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Umur"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000007&
         Caption         =   "Label18"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5160
         Width           =   7815
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000007&
         Caption         =   "Label16"
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   7815
      End
      Begin VB.Label Label15 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   6480
         TabIndex        =   25
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pasien"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Satuan"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Obat"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dokter"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Dokter"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Resep"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   16
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
         Left            =   6120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
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
      TabIndex        =   31
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frm_resep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dataobat As Database
Dim rs_obat As Recordset

Sub kosong()
txtno_resep = ""
txtkd_dokter = ""
txtnama_dokter = ""
txtkd_pasien = ""
txtnama_pasien = ""
txtumur_Pasien = ""
txtkd_obat = ""
txtNama_obat = ""
txtjml_obat = ""
txtharga = ""
txtJumlah = ""
txtstockawal = ""
txtstockakhir = ""
Text1.Text = ""
Text2 = ""
Text3 = ""
Text4 = ""
txtTotal = ""
End Sub
Sub buatnoresep()
    If Data5.Recordset.RecordCount <> 0 Then
       Data5.Recordset.MoveLast
       txtno_resep = Data5.Recordset!No_resep
       noint = Right(txtno_resep, 3) + 1
       norsp = Format(Date, "yymm")
       Text4.Text = "R" & norsp
       Select Case noint
           Case 0 To 9
               txtno_resep.Text = Text4 & "00" & (Trim(Str(noint)))
           Case 10 To 99
               txtno_resep.Text = Text4 & "0" & Trim(Str(noint))
           Case 100 To 999
               txtno_resep.Text = Text4 & Trim(Str(noint))
       End Select
    Else
       norsp = Format(Date, "yymm")
       txtno_resep.Text = "R" & norsp & "001"
    End If
    
    'norsp = Format(Date, "yymm")
    'If Data5.Recordset.RecordCount <> 0 Then
    '   Data5.Recordset.MoveLast
    '   txtno_resep = Right(Data5.Recordset!No_resep, 3)
    '   norsp = "R" & norsp & (Val(txt_noresep) + 1)
    'Else
    '  Data5.Recordset.MoveLast
    '  If Val(Mid(Data5.Recordset!No_resep, 3, 4)) <> Format(Date, "yymm") Then
    '     norsp = "R" & Format(Date, "yymm") & "001"
    '  Else
     '   noint = Val(Right(Data5.Recordset!No_resep, 3)) + 1
     '   norsp = "R" & Format(Date, "yymm") & Right("000" & noint, 3)
     ' End If
    'End If
     '   txtno_resep.Text = norsp
End Sub

Private Sub cmd_Hapus_Click()
If Data4.Recordset.RecordCount = 0 Then
   MsgBox "Maaf Tidak ada Yang Harus dibatalkan", vbInformation, "Info"
   Grid.Refresh
   Exit Sub
End If
x = MsgBox("Betul data akan dihapus ?", vbOKCancel + vbInformation, "Menghapus Record")
    If x = vbOK Then
        
         Data3.Recordset.Index = "Kode_obat"
            Data3.Recordset.Seek "=", Text1.Text
            If Not Data3.Recordset.NoMatch Then
                Data3.Recordset.edit
                Data3.Recordset!jml_obat = Val(Data3.Recordset!jml_obat) + Val(Text2.Text)
                txtTotal.Text = Val(txtTotal.Text) - Val(Text3.Text)
                Data3.Recordset.Update
            End If
         Data4.Recordset.delete
         Data4.Recordset.MoveFirst
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
If txtno_resep = "" Or txtkd_pasien = "" Or Data4.Recordset.RecordCount = 0 Then
   MsgBox "maaf data harus lengkap"
   Exit Sub
Else
  If Data4.Recordset.RecordCount > 0 Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
Data5.Recordset.AddNew
Data5.Recordset!No_resep = txtno_resep
Data5.Recordset!tgl = txtTanggal.Text
Data5.Recordset!kode_dokter = txtkd_dokter.Text
Data5.Recordset!nama_dokter = txtnama_dokter.Text
Data5.Recordset!kode_pasien = txtkd_pasien.Text
Data5.Recordset!nama_Pasien = txtnama_pasien.Text
Data5.Recordset!umur = txtumur_Pasien.Text
Data5.Recordset!kode_obat = Data4.Recordset!kode_obat
Data5.Recordset!nama_obat = Data4.Recordset!nama_obat
Data5.Recordset!jml_obat = Data4.Recordset!jml_obat
Data5.Recordset!harga_satuan = Data4.Recordset!harga_obat
Data5.Recordset!jumlah = Data4.Recordset!jumlah
Data5.Recordset.Update
            Data4.Recordset.MoveNext
        Loop
    End If
    If Data4.Recordset.RecordCount > 0 Then
        Data4.Recordset.MoveFirst
        Do While Not Data4.Recordset.EOF
            Data4.Recordset.delete
            Data4.Recordset.MoveNext
        Loop
    End If
Call kosong
Command1.Enabled = True
cmd_Simpan.Enabled = False
End If
End Sub

Private Sub cmd_tambah_Click()
If txtkd_obat.Text = "" Then
   MsgBox "isi dulu kode obatnya", vbInformation, "info"
   txtkd_obat.SetFocus
   Exit Sub
End If
   Data4.Recordset.AddNew
   Data4.Recordset!kode_obat = txtkd_obat.Text
   Data4.Recordset!nama_obat = txtNama_obat.Text
   Data4.Recordset!jml_obat = txtjml_obat.Text
   Data4.Recordset!harga_obat = txtharga.Text
   Data4.Recordset!jumlah = txtJumlah.Text
   Data4.Recordset.Update
   
            Data3.Recordset.Index = "Kode_obat"
            Data3.Recordset.Seek "=", txtkd_obat
          
            If Not Data3.Recordset.NoMatch Then
                Data3.Recordset.edit
                Data3.Recordset!jml_obat = Val(Data3.Recordset!jml_obat) - Val(txtjml_obat.Text)
                Data3.Recordset.Update
            End If
            txtTotal = Val(txtTotal.Text) + Val(txtJumlah.Text)
   txtkd_obat.Text = ""
   txtNama_obat.Text = ""
   txtjml_obat.Text = ""
   txtharga.Text = ""
   txtJumlah.Text = ""
End Sub

Private Sub cmdSalinanResep_Click()
frm_ctk_salinan_resep.Show
End Sub

Private Sub Command1_Click()

Call buatnoresep
txtTanggal = Format(Date, "dd-mm-yyyy")
txtkd_dokter.SetFocus
If Data4.Recordset.RecordCount > 0 Then
    Data4.Recordset.MoveFirst
    Do While Not Data4.Recordset.EOF
       Data4.Recordset.delete
       Data4.Recordset.MoveNext
    Loop
 End If
 Command1.Enabled = False
 cmd_Simpan.Enabled = True
End Sub

Private Sub Form_Activate()
If Text5.Text = "" Then Exit Sub:

End Sub

Private Sub Form_Load()
Command1.Enabled = True
End Sub

Private Sub Grid_Click()
Text1.Text = Grid.Columns(0)
Text2.Text = Grid.Columns(2)
Text3.Text = Grid.Columns(4)
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

Private Sub txtkd_dokter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_browse_dokter.Show
frm_browse_dokter.Grid.Refresh
End If
End Sub

Private Sub txtkd_obat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frm_browse_obat.Show
frm_browse_obat.Grid.Refresh
End If
End Sub

Private Sub txtkd_pasien_KeyPress(KeyAscii As Integer)
frm_browse_pasien.Show
frm_browse_pasien.Grid.Refresh
End Sub

