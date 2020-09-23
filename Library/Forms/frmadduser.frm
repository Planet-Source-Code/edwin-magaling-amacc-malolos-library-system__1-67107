VERSION 5.00
Begin VB.Form frmadduser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding Username and Password"
   ClientHeight    =   2115
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4575
   Icon            =   "frmadduser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1249.612
   ScaleMode       =   0  'User
   ScaleWidth      =   4295.677
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVerifyPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   900
      Width           =   2355
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1890
      TabIndex        =   0
      Top             =   90
      Width           =   2355
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   390
      Left            =   945
      TabIndex        =   3
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&CANCEL"
      Height          =   390
      Left            =   2595
      TabIndex        =   4
      Top             =   1575
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   495
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password:"
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
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   990
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   630
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Height          =   195
      Left            =   765
      TabIndex        =   5
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmadduser"
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

Private Sub cmdSave_Click()
If txtUserName = Empty Then
    MsgBox "Please Input New Username!", vbOKOnly + vbExclamation, "Information"
    txtUserName.SetFocus
    Exit Sub
End If
If txtPassword = Empty Then
    MsgBox "Please Input New Password!", vbOKOnly + vbExclamation, "Information"
    txtPassword.SetFocus
    Exit Sub
End If
If txtVerifyPassword = Empty Then
    MsgBox "Please Input to Verify Password!", vbOKOnly + vbExclamation, "Information"
    txtVerifyPassword.SetFocus
    Exit Sub
End If
If MsgBox("Save New Username and Password?", vbYesNo + vbQuestion, "Saving Username and Password") = vbNo Then
    Exit Sub
End If
    
    'users
    Set userRS = New ADODB.Recordset
    SQLstr = "select username from userlist where username='" & Trim(txtUserName.Text) & "'"
    userRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
    If userRS.BOF And userRS.EOF Then
    
        If txtPassword = txtVerifyPassword Then
            Set userRS = New ADODB.Recordset
            userRS.Open "userlist", libCON, adOpenKeyset, adLockOptimistic
            With userRS
                .AddNew
                !Username = txtUserName.Text
                !Password = txtPassword.Text
                .Update
                .Close
            End With
    
            txtUserName.Text = ""
            txtPassword.Text = ""
            txtVerifyPassword.Text = ""
            txtUserName.SetFocus
        Else
            MsgBox "Verify Password"
            txtVerifyPassword.Text = ""
            txtVerifyPassword.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Username Already Exists!", vbOKOnly + vbInformation, "Information"
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtVerifyPassword.Text = ""
        txtUserName.SetFocus
        Exit Sub
    End If
       
    frmusers.List1.clear

    'users
    Set userRS = New ADODB.Recordset
    userRS.Open "userlist", libCON, adOpenKeyset, adLockReadOnly

    While userRS.EOF <> True
        frmusers.List1.AddItem userRS!Username
        userRS.MoveNext
    Wend


frmusers.cmdDelete.Enabled = False
frmusers.cmdchange.Enabled = False

MsgBox "Username and Password Saved!", vbOKOnly + vbInformation, "Information"
Unload Me

End Sub

Private Sub Form_Load()
  dbconnect
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVerifyPassword.SetFocus
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub

Private Sub txtVerifyPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
