VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_dokter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DOKTER"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "dokter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8895
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "dokter.frx":000C
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "dokter.frx":0020
      TabIndex        =   23
      Top             =   360
      Width           =   5775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Dokter"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtkd_dok 
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
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtnm_dok 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Simpan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Isi Data"
      Height          =   495
      Left            =   6600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Batal 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ba&tal"
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      Caption         =   "Label10"
      Height          =   735
      Index           =   0
      Left            =   6240
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label4"
      Height          =   2895
      Index           =   1
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Label7"
      Height          =   2895
      Left            =   8640
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   8775
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Label4"
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Dokter"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Dokter"
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000C&
      Height          =   135
      Index           =   1
      Left            =   6240
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Caption         =   "Label9"
      Height          =   3375
      Index           =   0
      Left            =   5880
      TabIndex        =   9
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Label8"
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   19
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      Height          =   3375
      Index           =   1
      Left            =   6000
      TabIndex        =   20
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      Height          =   3375
      Index           =   2
      Left            =   8760
      TabIndex        =   21
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000C&
      Height          =   2775
      Left            =   6120
      TabIndex        =   22
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "frm_dokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bisa()
txtkd_dok.Enabled = True
txtnm_dok.Enabled = True
End Sub
Private Sub tdk_bisa()
txtkd_dok.Enabled = False
txtnm_dok.Enabled = False
tombol_normal
End Sub
Private Sub kunci()
txtkd_dok.Locked = True
txtnm_dok.Locked = True
End Sub

Private Sub bersih()
txtkd_dok.Text = ""
txtnm_dok.Text = ""
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

Sub tampil()
txtkd_dok.Text = Data1.Recordset!kode_dokter
txtnm_dok.Text = Data1.Recordset!nama_dokter
End Sub

Private Sub nomor()
Dim urutan As String * 5
Dim hitung As Byte

    If Data1.Recordset.RecordCount = 0 Then
        urutan = "Dr" & "001"
    Else
        Data1.Recordset.MoveLast
        If Val(Left(Data1.Recordset!kode_dokter, 3)) <> "000" Then
            urutan = "00" & "001"
        Else
        hitung = Val(Right(Data1.Recordset!kode_dokter, 3)) + 1
        urutan = "Dr" & Right("000" & hitung, 3)
    End If
    End If
    Me.txtkd_dok = urutan

End Sub

Private Sub cmd_Batal_Click()
Call bersih
Call tdk_bisa
cmd_Batal.Enabled = False
cmd_Perbaiki.Enabled = True
cmd_Hapus.Enabled = True
cmd_cari.Enabled = True
End Sub

Private Sub cmd_cari_Click()
Dim var As String
var = InputBox("Masukan Kode Dokter yang ingin anda cari!", "Cari data dokter")
If var = Empty Then Exit Sub
   If var <> "" Then
      Data1.Recordset.Index = "kode_dokter"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Call tampil
         Call bisa
         Call kunci
      Else
         MsgBox "Data dokter dengan kode dokter " & var & " tidak diketemukan"
      End If
   End If
End Sub

Private Sub cmd_Hapus_Click()
Dim var As String
var = InputBox("Masukan Kode dokter yang akan dihapus!", "Hapus dokter")
If var = Empty Then Exit Sub
   If var <> "" Then
      Data1.Recordset.Index = "Kode_dokter"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Data1.Recordset.Delete
         Data1.Refresh
         Data1.Recordset.MoveFirst
      Else
         MsgBox "Data dokter dengan kode dokter " & var & " tidak diketemukan"

      End If
    End If
End Sub

Private Sub cmd_Keluar_Click()
Unload Me
MDIForm1.dokter.Enabled = True
MDIForm1.dokter.Checked = False
End Sub

Private Sub cmd_Perbaiki_Click()
If cmd_Perbaiki.Caption = "Pe&rbaiki" Then
   cmd_Simpan.Enabled = False
   cmd_Hapus.Enabled = False
   cmd_Batal.Enabled = True
   Dim var As String
   var = InputBox("Ketikkan kode dokter yang datanya akan di perbaiki !", "Perbaiki Data dokter")
   If var = Empty Then Exit Sub
      Data1.Recordset.Index = "Kode_dokter"
      Data1.Recordset.Seek "=", var
      If Not Data1.Recordset.NoMatch Then
         Call tampil
         txtkd_dok.Enabled = False
         txtnm_dok.Enabled = True
         cmd_Perbaiki.Caption = "&Perbaharui data"
      Else
         MsgBox "Data dokter dengan kode dokter " & var & " tidak diketemukan"
      End If
Else
Data1.Recordset.Edit
Data1.Recordset!kode_dokter = txtkd_dok.Text
Data1.Recordset!nama_dokter = txtnm_dok.Text
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
nomor
Me.txtnm_dok.SetFocus
cmd_Batal.Enabled = True
cmd_Perbaiki.Enabled = False
cmd_Hapus.Enabled = False
cmd_cari.Enabled = False
cmd_Simpan.Caption = "&Simpan Data"
Else
If txtkd_dok.Text = "" Or _
        txtnm_dok.Text = "" Then
        MsgBox "Data tidak boleh kosong !", vbCritical, "SISTEM PENJUALAN KREDIT"
        txtkd_dok.SetFocus
        Else
cmd_Batal.Enabled = False
cmd_Perbaiki.Enabled = True
cmd_Hapus.Enabled = True
cmd_cari.Enabled = True
Data1.Recordset.AddNew
Data1.Recordset!kode_dokter = txtkd_dok.Text
Data1.Recordset!nama_dokter = txtnm_dok.Text
Data1.Recordset.Update
Call bersih
cmd_Simpan.Caption = "&Isi Data"
End If
End If
End Sub

