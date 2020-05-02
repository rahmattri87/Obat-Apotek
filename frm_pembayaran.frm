VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_pembayaran 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBAYARAN"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Pembayaran"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Resep"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Height          =   735
         Left            =   4200
         TabIndex        =   15
         Top             =   4320
         Width           =   2775
         Begin VB.CommandButton cmdsimpan 
            Caption         =   "&Simpan"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdtutup 
            Caption         =   "&Tutup"
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdbatal 
            Caption         =   "&Batal"
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtukem 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtubay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txttobay 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   3960
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frm_pembayaran.frx":0000
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "frm_pembayaran.frx":0014
         TabIndex        =   8
         Top             =   960
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6855
         Begin VB.TextBox txtno_bayar 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txttgl_byr 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtnoresep 
            Height          =   285
            Left            =   960
            MaxLength       =   8
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label txtnobyr 
            BackStyle       =   0  'Transparent
            Caption         =   "No Bayar"
            Height          =   255
            Left            =   4800
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Bayar"
            Height          =   255
            Left            =   2280
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No Resep"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Kembali"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Uang Bayar"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bayar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   0
         X2              =   7080
         Y1              =   3840
         Y2              =   3840
      End
   End
End
Attribute VB_Name = "frm_pembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub buatnobayar()
  If Data2.Recordset.RecordCount <> 0 Then
       Data2.Recordset.MoveLast
       txtno_bayar = Data2.Recordset!no_bayar
       noint = Right(txtno_bayar, 3) + 1
       nobyr = Format(Date, "yymm")
       Text4.Text = "B" & nobyr
       Select Case noint
           Case 0 To 9
               txtno_bayar.Text = Text4 & "00" & (Trim(Str(noint)))
           Case 10 To 99
               txtno_bayar.Text = Text4 & "0" & Trim(Str(noint))
           Case 100 To 999
               txtno_bayar.Text = Text4 & Trim(Str(noint))
       End Select
    Else
       nobyr = Format(Date, "yymm")
       Text4.Text = "B" & nobyr
       txtno_bayar.Text = Text4 & "001"
    End If
End Sub

Sub kosong()
txtnoresep = ""
txttobay = ""
txtubay = ""
txtukem = ""
End Sub

Private Sub cmdbatal_Click()
Data1.RecordSource = "select * from resep "
Data1.Refresh
kosong
txtnoresep.Enabled = True
DBGrid1.Visible = False
End Sub

Private Sub cmdsimpan_Click()
Data1.RecordSource = "select * from resep where no_resep = '" & txtnoresep.Text & "'"
        Data1.Refresh
        If Data1.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Maaf No Resep " & txtnoresep & " belum ada dalam database kami", vbInformation, "Perhatian"
            txtnoresep.SetFocus
            txtnoresep.SelStart = 0
            txtnoresep.SelLength = Len(txtnoresep.Text)
            Exit Sub
        Else
        Data1.RecordSource = "select * from resep where no_resep = '" & txtnoresep.Text & "'"
        Data1.Refresh
        Do While Not Data1.Recordset.EOF
        'Data1.Recordset.MoveFirst
       
        Data2.Recordset.AddNew
Data2.Recordset!no_bayar = txtno_bayar
Data2.Recordset!tgl_byr = txttgl_byr.Text
Data2.Recordset!tgl_resep = Data1.Recordset!tgl
Data2.Recordset!No_resep = Data1.Recordset!No_resep
Data2.Recordset!kode_dokter = Data1.Recordset!kode_dokter
Data2.Recordset!nama_dokter = Data1.Recordset!nama_dokter
Data2.Recordset!kode_pasien = Data1.Recordset!kode_pasien
Data2.Recordset!nama_Pasien = Data1.Recordset!nama_Pasien
Data2.Recordset!umur = Data1.Recordset!umur
Data2.Recordset!kode_obat = Data1.Recordset!kode_obat
Data2.Recordset!nama_obat = Data1.Recordset!nama_obat
Data2.Recordset!jml_obat = Data1.Recordset!jml_obat
Data2.Recordset!harga_satuan = Data1.Recordset!harga_satuan
Data2.Recordset!jumlah = Data1.Recordset!jumlah
Data2.Recordset.Update
            'Data2.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
    End If
    
     x = MsgBox("mau cetak kwitansi pembayaran?", vbInformation + vbYesNo, "Info")
 If x = vbYes Then
  frm_ctk_Kwitansi_pembayaran.Show
 Else
    kosong
    buatnobayar
    Data1.RecordSource = "select * from resep "
    Data1.Refresh
    DBGrid1.Visible = False
    cmdsimpan.Enabled = False
    txtnoresep.Enabled = True
  End If
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txttgl_byr = Format(Date, "DD-MM-YYYY")
Call kosong
Call buatnobayar
DBGrid1.Visible = False
End Sub

Private Sub txtnoresep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.Visible = True
   DBGrid1.SetFocus
End If
End Sub

Private Sub txtnoresep_LostFocus()
If txtnoresep.Text = "" Then Exit Sub:
        Data1.RecordSource = "select * from resep where no_resep = '" & txtnoresep.Text & "'"
        Data1.Refresh
        If Data1.Recordset.RecordCount = 0 Then
            Beep
            MsgBox "Maaf No Resep " & txtnoresep & " belum ada dalam database kami", vbInformation, "Perhatian"
            txtnoresep.SetFocus
            txtnoresep.SelStart = 0
            txtnoresep.SelLength = Len(txtnoresep.Text)
            DBGrid1.Visible = False
            Exit Sub
        Else
            
            txtnoresep.Enabled = False
            DBGrid1.Visible = True
            If Not Data1.Recordset.RecordCount = 0 Then
                   Data1.Recordset.MoveFirst
                   Do While Not Data1.Recordset.EOF
                   hit = hit + Data1.Recordset!jumlah
                   Data1.Recordset.MoveNext
                   Loop
            End If

            txttobay = hit
            DBGrid1.SetFocus
       End If
End Sub

Private Sub txtubay_Change()
txtukem.Text = Val(txtubay.Text) - Val(txttobay.Text)
If txtukem.Text >= 0 Then
   cmdsimpan.Enabled = True
Else
   cmdsimpan.Enabled = False
End If

End Sub
