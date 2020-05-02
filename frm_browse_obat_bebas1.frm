VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_browse_obat_bebas 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BROWSE OBAT BEBAS"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid Grid 
      Bindings        =   "frm_browse_obat_bebas1.frx":0000
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frm_browse_obat_bebas1.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
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
      RecordSource    =   "Obat"
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "frm_browse_obat_bebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Grid_DblClick()
frm_obat_bebas.txtkd_obat = Grid.Columns(0).Text
frm_obat_bebas.txtNama_obat = Grid.Columns(1).Text
frm_obat_bebas.txtharga = Grid.Columns(3).Text
frm_obat_bebas.txtstockawal = Grid.Columns(4).Text
frm_obat_bebas.Text5 = Grid.Columns(5).Text
frm_obat_bebas.txtjml_obat.SetFocus
Unload Me
End Sub
