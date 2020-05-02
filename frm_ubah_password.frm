VERSION 5.00
Begin VB.Form frmubahpassword 
   Caption         =   "Form Password"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      MouseIcon       =   "frm_ubah_password.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      MouseIcon       =   "frm_ubah_password.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frm_ubah_password.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frm_ubah_password.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtnama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtuserid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtpasswordbaru 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MouseIcon       =   "frm_ubah_password.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Baru "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1890
   End
End
Attribute VB_Name = "frmubahpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rspegawai As Recordset
Dim edit As Boolean
Dim delete As Boolean

Private Sub cmdadd_Click()
 If cmdadd.Caption = "Add" Then
  cmdadd.Caption = "Cancel"
    cmdok.Caption = "Save"
    txtnama.Visible = True
  'txtpasswordbaru.Visible = True
  'Label3.Visible = True
   Label4.Visible = True
    txtnama.SetFocus
    edit = False
  Else
    txtpasswordbaru.Visible = False
    Label3.Visible = False
    txtnama.Visible = False
    Label4.Visible = False
    cmdok.Caption = "Login"
    cmdadd.Caption = "Add"
  End If
End Sub

Private Sub cmddelete_Click()
delete = True
cmdok.Caption = "Save"
txtuserid.SetFocus
cmdadd.Caption = "Cancel"
End Sub

Private Sub cmdedit_Click()
edit = True
txtuserid.SetFocus
cmdok.Caption = "Save"
cmdadd.Caption = "Cancel"
  txtpasswordbaru.Visible = True
  Label3.Visible = True
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
If cmdok.Caption = "Login" Then
  Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai='" & txtuserid & "'")

  If Trim(txtuserid.Text) = rspegawai!idpegawai And txtpassword.Text = Trim(rspegawai!Password) Then
  MDIForm1.Show
  Unload Me
  Else
  MsgBox "Wrong username or password!", vbCritical, "Warning"
  End If
Else
If edit = True Then
Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai like '" & txtuserid & "*'")
If Not rspegawai.NoMatch Then

If Trim(txtuserid.Text) = rspegawai("idpegawai") And Trim(txtpassword.Text) = Trim(rspegawai("password")) Then
rspegawai.edit
rspegawai("Password") = txtpasswordbaru.Text
rspegawai("idpegawai") = txtuserid.Text
rspegawai.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Userid = " & txtuserid.Text & Chr(13) & "Password = " & txtpasswordbaru.Text, vbInformation, "Warning")
MDIForm1.Show
Unload Me
Else
MsgBox "Passwords don't match!", vbCritical, "Warning!"
txtpassword.Text = vbNullString
txtpasswordbaru.Text = vbNullString
End If
End If
Else
rspegawai.AddNew
rspegawai("nama") = txtnama.Text
rspegawai("Password") = txtpassword.Text
rspegawai("idpegawai") = txtuserid.Text
rspegawai.Update
End If
cmdadd.Caption = "Add"
cmdok.Caption = "Login"
txtuserid = ""
txtpassword = ""
  txtnama.Visible = False
  txtpasswordbaru.Visible = False
  Label3.Visible = False
  Label4.Visible = False
edit = False
End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Activate()
  cmdok.Caption = "Login"
  txtnama.Visible = False
  txtpasswordbaru.Visible = False
  Label3.Visible = False
  Label4.Visible = False
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path & "\database\mydb.mdb")
Set rspegawai = db.OpenRecordset("pegawai", dbOpenDynaset)
End Sub



Private Sub txtuserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmdok.Caption = "Login" Then
  Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai='" & txtuserid & "'")
  If Trim(txtuserid.Text) = rspegawai!idpegawai And txtpassword.Text = Trim(rspegawai!Password) Then
  MDIForm1.Show
  Unload Me
  Else
  MsgBox "Wrong username or password!", vbCritical, "Warning"
  End If
Else
Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai like '" & txtuserid & "*'")
If delete = False Then
    If edit = False Then
       If Not rspegawai.EOF Then
            MsgBox "Data Sudah ada", 0, "Info"
            txtuserid.SetFocus
        Else
            txtnama.SetFocus
        End If
    Else
        If Not rspegawai.EOF Then
          txtnama.Text = rspegawai("nama")
     
          txtpasswordbaru.Enabled = True
        Else
            MsgBox "Data tidak ada", 0, "Info"
            txtuserid.SetFocus
        End If
        End If
Else
Set rspegawai = db.OpenRecordset("select * from pegawai where idpegawai like '" & txtuserid & "*'")
Dim x As Integer
x = MsgBox("Yakin mau di hapus", vbOKCancel, "Info")
If x = 1 Then
If Not rspegawai.EOF Then

   rspegawai.delete
End If
cmdok.Caption = "Login"
cmdadd.Caption = "Add"
txtuserid = ""
delete = False
Else
cmdok.Caption = "Login"
cmdadd.Caption = "Add"
txtuserid = ""
delete = False
End If
End If
End If
End If


End Sub

