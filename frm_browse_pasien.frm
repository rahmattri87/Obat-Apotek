VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_browse_pasien 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Pasien"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "frm_browse_pasien.frx":0000
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frm_browse_pasien.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Aplikasi _Penjualan_Apotik_ASTITI\Database\MyDB.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Pasien"
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "frm_browse_pasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_DblClick()
frm_resep.txtkd_pasien = Grid.Columns(0).Text
frm_resep.txtnama_Pasien = Grid.Columns(1).Text
frm_resep.txtumur_Pasien = Grid.Columns(2).Text
frm_resep.txtkd_obat.SetFocus
Unload Me
End Sub
