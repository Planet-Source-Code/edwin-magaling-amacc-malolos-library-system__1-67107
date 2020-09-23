VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AMA - Library System"
   ClientHeight    =   1800
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4560
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1063.5
   ScaleMode       =   0  'User
   ScaleWidth      =   4281.593
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   615
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   390
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Log-In"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&CANCEL"
      Height          =   390
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   1260
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   180
      Picture         =   "frmlogin.frx":0442
      Stretch         =   -1  'True
      Top             =   135
      Width           =   510
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Name :"
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
      Height          =   270
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password :"
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
      Height          =   270
      Index           =   1
      Left            =   990
      TabIndex        =   4
      Top             =   675
      Width           =   945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pwctr As Integer

Private Sub cmdCancel_Click()
ReleaseMenus hwnd
End
End Sub

Private Sub cmdOK_Click()

Set userRS = New ADODB.Recordset
If txtUserName.Text <> "" Then
    
    SQLstr = "Select * From userlist Where username = '" & Trim(txtUserName.Text) & "'"
    userRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If txtUserName.Text <> "Administrator" Then
    If txtUserName.Text <> "administrator" Then
    If txtUserName.Text <> "ADMINISTRATOR" Then
        frmmain.mnusettingsystem.Enabled = False
        frmmain.mnusettinguser.Enabled = False
    Else
        frmmain.mnusettingsystem.Enabled = True
        frmmain.mnusettinguser.Enabled = True
    End If
    Else
        frmmain.mnusettingsystem.Enabled = True
        frmmain.mnusettinguser.Enabled = True
    End If
    Else
        frmmain.mnusettingsystem.Enabled = True
        frmmain.mnusettinguser.Enabled = True
    End If
    If Not userRS.EOF And Not userRS.BOF Then
        If txtPassword.Text <> userRS!Password Then
            pwctr = pwctr + 1
            If pwctr = 1 Then
                MsgBox "Invalid password! You have 2 tries remaining!", vbOKOnly + vbInformation, "Information"
                txtPassword.Text = ""
                txtPassword.SetFocus
            
            ElseIf pwctr = 2 Then
                MsgBox "Invalid password! You only have 1 try remaining!", vbOKOnly + vbInformation, "Information"
                txtPassword.Text = ""
                txtPassword.SetFocus
            Else
               ReleaseMenus hwnd
               End
            End If
        Else
            Unload Me
            frmmain.Show
        End If
    Else
        MsgBox "Invalid Username!", vbOKOnly + vbExclamation, "Warning.."
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtUserName.SetFocus
    End If
Else
    MsgBox "Invalid Username and Password!", vbOKOnly + vbExclamation, "Warning.."
    txtUserName.SetFocus
End If

End Sub

Private Sub Form_Load()
    dbconnect
    txtUserName.Text = GetSetting(App.EXEName, "TextBox", txtUserName.Name, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Textbox", txtUserName.Name, txtUserName.Text
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdOK.Enabled = True Then
        cmdOK.Value = True
    End If
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub
