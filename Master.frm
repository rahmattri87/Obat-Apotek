VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistem Informasi Penjualan "
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10770
   Icon            =   "Master.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Master.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7200
      Top             =   3240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   15460
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Visible         =   0   'False
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            TextSave        =   "2:53 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&Master"
      Begin VB.Menu obat 
         Caption         =   "Obat"
      End
      Begin VB.Menu pasien 
         Caption         =   "Pasien"
      End
      Begin VB.Menu dokter 
         Caption         =   "Dokter"
      End
      Begin VB.Menu resep 
         Caption         =   "Resep"
      End
   End
   Begin VB.Menu mnujual 
      Caption         =   "&Pembayaran"
      Begin VB.Menu kwitansai 
         Caption         =   "Kwitansi"
      End
      Begin VB.Menu mnuObatBebas 
         Caption         =   "Obat Bebas"
      End
   End
   Begin VB.Menu laporan 
      Caption         =   "&Buat Laporan"
      Begin VB.Menu lapObatResep 
         Caption         =   "Laporan Penjualan Obat Resep"
      End
      Begin VB.Menu lapObatBebas 
         Caption         =   "Laporan Penjualan Obat Bebas"
      End
   End
   Begin VB.Menu Mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dokter_Click()
frm_dokter.Show
dokter.Checked = True
dokter.Enabled = False

'Dim RsObatResep As New ADODB.Recordset
End Sub

Private Sub kwitansai_Click()
frm_pembayaran.Show
End Sub

Private Sub lapObatBebas_Click()
'Dim con As New ADODB.Connection
'Dim RsObatBebas As New ADODB.Recordset
'con.Open "provider=microsoft.jet.oledb.4.0;data source= " & App.Path & "\Database\mydb.mdb"
'RsObatBebas.Open "SELECT struk_obat_bebas.no_struk, struk_obat_bebas_detail.kode_obat, obat.nama_obat, obat.satuan, struk_obat_bebas_detail.jml_obat, struk_obat_bebas.tgl_struk FROM obat, struk_obat_bebas INNER JOIN struk_obat_bebas_detail ON struk_obat_bebas.no_struk = struk_obat_bebas_detail.no_struk order by struk_obat_bebas.no_struk", con, adOpenKeyset, adLockOptimistic
'Set DataReport3.DataSource = RsObatBebas
'DataReport3.Show
lap_obat_bebas.Show
End Sub

Private Sub lapObatResep_Click()
'Dim con As New ADODB.Connection
'Dim RsObatResep As New ADODB.Recordset
'con.Open "provider=microsoft.jet.oledb.4.0;data source= " & App.Path & "\Database\mydb.mdb"
'RsObatResep.Open "SELECT resep_detail.no_resep, resep_detail.kode_obat, obat.nama_obat, obat.satuan, resep_detail.jml_obat, resep.tgl_resep FROM (obat INNER JOIN resep_detail ON obat.kode_obat = resep_detail.kode_obat) INNER JOIN resep ON resep_detail.no_resep = resep.no_resep order by resep_detail.no_resep", con, adOpenKeyset, adLockOptimistic
'Set DataReport2.DataSource = RsObatResep
'DataReport2.Show
lap_data_pembayaran.Show
End Sub





Private Sub Mnuexit_Click()
Dim tanya As String
tanya = MsgBox("Apakah anda ingin keluar ???", vbYesNo + vbInformation, "Pesan")
If tanya = vbYes Then
Unload Me
Else
Exit Sub
End If
End Sub

Private Sub mnuObatBebas_Click()
frm_obat_bebas.Show
End Sub

Private Sub mnuObatResep_Click()
frm_resep.Show
End Sub

Private Sub obat_Click()
frm_obat.Show
obat.Checked = True
obat.Enabled = False
End Sub
Private Sub pasien_Click()
frm_pasien.Show
pasien.Checked = True
pasien.Enabled = False
End Sub

Private Sub resep_Click()
frm_resep.Show
End Sub
