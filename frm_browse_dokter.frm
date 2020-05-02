VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_browse_dokter 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Dokter"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Dokter"
      Top             =   4440
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "frm_browse_dokter.frx":0000
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frm_browse_dokter.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frm_browse_dokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_DblClick()
frm_resep.txtkd_dokter = Grid.Columns(0).Text
frm_resep.txtnama_dokter = Grid.Columns(1).Text
frm_resep.txtkd_pasien.SetFocus
Unload Me
End Sub

