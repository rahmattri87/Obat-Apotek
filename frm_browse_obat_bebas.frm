VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_browse_obat 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Obat"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
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
      RecordsetType   =   0  'Table
      RecordSource    =   "Obat"
      Top             =   4320
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "frm_browse_obat_bebas.frx":0000
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "frm_browse_obat_bebas.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frm_browse_obat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Grid_DblClick()
frm_resep.txtkd_obat = Grid.Columns(0).Text
frm_resep.txtNama_obat = Grid.Columns(1).Text
frm_resep.txtharga = Grid.Columns(3).Text
frm_resep.txtstockawal = Grid.Columns(4).Text
frm_resep.txtjml_obat.SetFocus
Unload Me
End Sub
