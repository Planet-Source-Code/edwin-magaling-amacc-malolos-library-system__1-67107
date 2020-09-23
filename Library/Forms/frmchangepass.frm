VERSION 5.00
Begin VB.Form frmchange 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Changing Password"
   ClientHeight    =   2115
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4470
   Icon            =   "frmchangepass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1249.612
   ScaleMode       =   0  'User
   ScaleWidth      =   4197.087
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtoldpassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   450
      Width           =   2355
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   945
      Width           =   2355
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "C&HANGE"
      Enabled         =   0   'False
      Height          =   390
      Left            =   720
      TabIndex        =   2
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&CANCEL"
      Height          =   390
      Left            =   2700
      TabIndex        =   4
      Top             =   1575
      Width           =   1140
   End
   Begin VB.Label lblusername 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1890
      TabIndex        =   7
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label Username 
      BackStyle       =   0  'Transparent
      Caption         =   "Username: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   855
      TabIndex        =   6
      Top             =   135
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   540
      TabIndex        =   5
      Top             =   540
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   990
      Width           =   1455
   End
End
Attribute VB_Name = "frmchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
With frmusers
    .List1.Text = ""
    .cmdDelete.Enabled = False
    .cmdchange.Enabled = False
End With
    Unload Me
End Sub

Private Sub cmdchange_Click()
If txtNewPassword = Empty Then
    MsgBox "Please input New Password!", vbOKOnly + vbInformation, "Information"
    Exit Sub
End If

If MsgBox("Change Password?", vbYesNo + vbQuestion, "Changing Password") = vbNo Then
    Exit Sub
End If
    
    'users
    Set userRS = New ADODB.Recordset
    SQLstr = "select * from userlist where username='" & Trim(lblusername.Caption) & "'"
    userRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
    With userRS
        !Password = txtNewPassword
        .Update
        .Close
    End With

With frmusers
    .List1.Text = ""
    .cmdchange.Enabled = False
    .cmdDelete.Enabled = False
End With

MsgBox "Password Changed!", vbOKOnly + vbInformation, "Information"
Unload Me

End Sub

Private Sub Form_Load()
   dbconnect
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdchange.SetFocus
End If
End Sub

Private Sub txtoldpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

'users
Set userRS = New ADODB.Recordset
SQLstr = "Select * From userlist Where username = '" & Trim(lblusername.Caption) & "'"
userRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    
    If Not userRS.EOF And Not userRS.BOF Then
        If txtoldpassword.Text <> userRS!Password Then
            MsgBox "Invalid Password", vbCritical, Caption
            txtoldpassword.Text = ""
            txtoldpassword.SetFocus
            Exit Sub
        Else
            txtNewPassword.Enabled = True
            txtNewPassword.SetFocus
            txtoldpassword.Enabled = False
            cmdchange.Enabled = True
        End If
    End If

End If
End Sub
