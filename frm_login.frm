VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtpassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtuserid 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   540
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rspegawai As Recordset
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdok_Click()
Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai='" & txtuserid & "'")

If Trim(txtuserid.Text) = rspegawai!idpegawai And txtpassword.Text = Trim(rspegawai!Password) Then
MDIForm1.Show
Unload Me
Else
MsgBox "Wrong username or password!", vbCritical, "Warning"
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\database\mydb.mdb")
Set rspegawai = db.OpenRecordset("pegawai", dbOpenDynaset)
End Sub



