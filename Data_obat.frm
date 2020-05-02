VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_obat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OBAT"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   8895
   Begin VB.TextBox txttgl_ex 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cbosatuan 
      Height          =   315
      Left            =   2160
      TabIndex        =   31
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txt_ins 
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   30
      Top             =   3360
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Obat"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Data_obat.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "Data_obat.frx":0014
      TabIndex        =   29
      Top             =   240
      Width           =   6015
   End
   Begin VB.TextBox txtkode_obat 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtnama 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtsatuan 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtJml_obat 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txthrg_satuan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Simpan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Isi Data"
      Height          =   495
      Left            =   6600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Batal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ba&tal"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Perbaiki 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pe&rbaiki"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Hapus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmd_cari 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cari Data"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Keluar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "K&eluar"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Expayerd"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   35
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "Label10"
      Height          =   735
      Index           =   0
      Left            =   6240
      TabIndex        =   25
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label4"
      Height          =   2895
      Index           =   1
      Left            =   6240
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Label7"
      Height          =   2895
      Left            =   8640
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   8775
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label4"
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Obat"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Obat"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Obat"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Satuan"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000C&
      Height          =   135
      Index           =   1
      Left            =   6240
      TabIndex        =   26
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      Height          =   3375
      Index           =   1
      Left            =   6000
      TabIndex        =   23
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      Height          =   3375
      Index           =   2
      Left            =   8760
      TabIndex        =   28
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Caption         =   "Label9"
      Height          =   3375
      Index           =   0
      Left            =   5880
      TabIndex        =   21
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Label8"
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000C&
      Height          =   2775
      Left            =   6120
      TabIndex        =   27
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "frm_obat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbosatuan_Click()
If cbosatuan.Text = "Tablet" Then
   Text1.Text = 1
   txtsatuan = "Tablet"
   nomor
ElseIf cbosatuan.Text = "Kapsul" Then
   Text1.Text = 2
    txtsatuan = "Kapsul"
    nomor
ElseIf cbosatuan.Text = "Cair" Then
   Text1.Text = 3
    txtsatuan = "Cair"
    nomor
ElseIf cbosatuan.Text = "Serbuk" Then
   Text1.Text = 4
    txtsatuan = "Serbuk"
    nomor
ElseIf cbosatuan.Text = "Salep" Then
   Text1.Text = 5
    txtsatuan = "Salep"
    nomor
End If
End Sub



Private Sub cmd_Batal_Click()
Call bersih
Call tdk_bisa
cmd_Batal.Enabled = False
cmd_Perbaiki.Enabled = True
cmd_Hapus.Enabled = True
cmd_cari.Enabled = True
End Sub

Private Sub nomor()
If Data1.Recordset.RecordCount <> 0 Then
       Data1.Recordset.MoveLast
       Text2.Text = Right(Data1.Recordset!kode_obat, 2) + 1
       noint = Text2 'Right(Text2, 2) + 1
       kdobt = txt_ins & Text1
       'Text4.Text = "B" & kdobt
       Select Case noint
           Case 0 To 9
               txtkode_obat.Text = kdobt & "0" & (Trim(Str(noint)))
           Case 10 To 99
               txtkode_obat.Text = kdobt & Trim(Str(noint))
           'Case 100 To 999
           '    txtno_bayar.Text = Text4 & Trim(Str(noint))
       End Select
    Else
       'nobyr = Format(Date, "yymm")
       'Text4.Text = "B" & nobyr
       txtkode_obat.Text = txt_ins & Text1 & "01"
    End If
'Dim urutan As String * 5
'Dim hitung As Byte
'    If Data1.Recordset.RecordCount = 0 Then
'        urutan = "OB" & "001"
'    Else
'        Data1.Recordset.MoveLast
'        If Val(Left(Data1.Recordset!kode_obat, 3)) <> "000" Then
'            urutan = "00" & "001"
'        Else
'        hitung = Val(Right(Data1.Recordset!kode_obat, 3)) + 1
'        urutan = "OB" & Right("000" & hitung, 3)
'    End If
'    End If
'    Me.txtkode_obat = urutan
End Sub
Private Sub cmd_cari_Click()
Dim var As String
var = InputBox("Masukan Kode obat yang ingin anda cari!", "Cari data obat")
If var = Empty Then Exit Sub
   If var <> "" Then
      Data1.Recordset.Index = "kode_obat"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Call tampil
         Call bisa
         Call kunci
      Else
         MsgBox "Data obat dengan kode obat " & var & " tidak diketemukan"
      End If
   End If
End Sub

Private Sub cmd_Hapus_Click()
Dim var As String
var = InputBox("Masukan Kode Obat yang akan dihapus!", "Hapus Obat")
If var = Empty Then Exit Sub
   If var <> "" Then
      Data1.Recordset.Index = "Kode_obat"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Data1.Recordset.delete
         Data1.Refresh
         Data1.Recordset.MoveFirst
      Else
         MsgBox "Data obat dengan kode obat " & var & " tidak diketemukan"

      End If
    End If
      
End Sub

Private Sub cmd_Keluar_Click()
Unload Me
MDIForm1.obat.Enabled = True
MDIForm1.obat.Checked = False
End Sub

Private Sub cmd_Perbaiki_Click()
If cmd_Perbaiki.Caption = "Pe&rbaiki" Then
   cmd_Simpan.Enabled = False
   cmd_Hapus.Enabled = False
   cmd_Batal.Enabled = True
   Dim var As String
   var = InputBox("Ketikkan kode Obat yang datanya akan di perbaiki !", "Perbaiki Data Obat")
   If var = Empty Then Exit Sub
      Data1.Recordset.Index = "Kode_Obat"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Call tampil
         Call bisa
         txtkode_obat.Enabled = False
         txtsatuan.Enabled = True
         txtnama.Enabled = True
         txthrg_satuan.Enabled = True
         txtjml_obat.Enabled = True
         cmd_Perbaiki.Caption = "&Perbaharui data"
      Else
         MsgBox "Data Obat dengan kode Obat " & var & " tidak diketemukan"
      End If
Else
Data1.Recordset.edit
Data1.Recordset!kode_obat = txtkode_obat
Data1.Recordset!nama_obat = txtnama
Data1.Recordset!satuan = txtsatuan
Data1.Recordset!jml_obat = txtjml_obat
Data1.Recordset!harga_satuan = txthrg_satuan
Data1.Recordset!tgl_ex = txttgl_ex.Text
Data1.Recordset.Update
Call bersih
cmd_Perbaiki.Caption = "Pe&rbaiki"
cmd_Batal.Enabled = False
cmd_Simpan.Enabled = True
cmd_Hapus.Enabled = True
Call tdk_bisa
End If
End Sub

Private Sub cmd_Simpan_Click()
If cmd_Simpan.Caption = "&Isi Data" Then
Call bisa
'nomor
Me.txt_ins.SetFocus
cmd_Batal.Enabled = True
cmd_Perbaiki.Enabled = False
cmd_Hapus.Enabled = False
cmd_cari.Enabled = False
cmd_Simpan.Caption = "&Simpan Data"
Else
If txtkode_obat.Text = "" Or _
        txtkode_obat.Text = "" Then
        MsgBox "Data tidak boleh kosong !", vbCritical, "SISTEM PENJUALAN KREDIT"
        txt_ins.SetFocus
        Else
cmd_Batal.Enabled = False
cmd_Perbaiki.Enabled = True
cmd_Hapus.Enabled = True
cmd_cari.Enabled = True
Data1.Recordset.AddNew
Data1.Recordset!kode_obat = txtkode_obat
Data1.Recordset!nama_obat = txtnama
Data1.Recordset!satuan = txtsatuan
Data1.Recordset!jml_obat = txtjml_obat
Data1.Recordset!harga_satuan = txthrg_satuan
Data1.Recordset!tgl_ex = txttgl_ex.Text
Data1.Recordset.Update
Call bersih
cmd_Simpan.Caption = "&Isi Data"
End If
End If

End Sub

Private Sub bisa()
txt_ins.Enabled = True
cbosatuan.Enabled = True
'txtkode_obat.Enabled = True
txtnama.Enabled = True
'txtsatuan.Enabled = True
txtjml_obat.Enabled = True
txthrg_satuan.Enabled = True
txttgl_ex.Enabled = True
End Sub
Private Sub tdk_bisa()
txt_ins.Enabled = False
cbosatuan.Enabled = False
txtkode_obat.Enabled = False
txtnama.Enabled = False
txtsatuan.Enabled = False
txtjml_obat.Enabled = False
txthrg_satuan.Enabled = False
txttgl_ex.Enabled = False
tombol_normal
End Sub
Private Sub kunci()
txt_ins.Locked = True
cbosatuan.Locked = True
txtkode_obat.Locked = True
txtnama.Locked = True
txtsatuan.Locked = True
txtjml_obat.Locked = True
txthrg_satuan.Locked = True
txttgl_ex.Locked = True
End Sub
Private Sub bersih()
txt_ins.Text = ""
cbosatuan.Text = ""
txtkode_obat.Text = ""
txtnama.Text = ""
txtsatuan.Text = ""
txtjml_obat.Text = ""
txthrg_satuan.Text = ""
txttgl_ex.Text = ""
End Sub
Private Sub tombol_normal()
cmd_Simpan.Caption = "&Isi Data"
cmd_Batal.Caption = "&Batal"
cmd_Perbaiki.Caption = "Pe&rbaiki"
cmd_Hapus.Caption = "&Hapus"
cmd_Simpan.Enabled = True
cmd_Batal.Enabled = False
cmd_Perbaiki.Enabled = True
cmd_Hapus.Enabled = True
End Sub
Private Sub tampil()

txtkode_obat = Data1.Recordset!kode_obat
txtnama = Data1.Recordset!nama_obat
txtsatuan = Data1.Recordset!satuan
txtjml_obat = Data1.Recordset!jml_obat
txthrg_satuan = Data1.Recordset!harga_satuan
txttgl_ex.Text = Data1.Recordset!tgl_ex
End Sub

Private Sub Form_Activate()
bersih
tdk_bisa
End Sub

Private Sub Form_Load()
cbosatuan.AddItem "Tablet"
cbosatuan.AddItem "Kapsul"
cbosatuan.AddItem "Cair"
cbosatuan.AddItem "Serbuk"
cbosatuan.AddItem "Salep"
End Sub
