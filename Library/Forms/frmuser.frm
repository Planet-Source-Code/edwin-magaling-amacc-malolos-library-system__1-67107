VERSION 5.00
Begin VB.Form frmusers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdchange 
      Caption         =   "CHANGE &PASSWORD"
      Enabled         =   0   'False
      Height          =   480
      Left            =   2250
      TabIndex        =   5
      Top             =   1395
      Width           =   1080
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CL&OSE"
      Height          =   435
      Left            =   2250
      TabIndex        =   4
      Top             =   2475
      Width           =   1080
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2250
      TabIndex        =   3
      Top             =   945
      Width           =   1080
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   435
      Left            =   2250
      TabIndex        =   1
      Top             =   495
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List of users:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   135
      Width           =   1680
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    frmadduser.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdchange_Click()

'users
Set userRS = New ADODB.Recordset
SQLstr = "SELECT * FROM userlist where username='" & List1.Text & "'"
userRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic

frmchange.lblusername.Caption = userRS!Username

frmchange.Show vbModal

End Sub

Private Sub cmdDelete_Click()

If List1.Text = "Administrator" Then
    MsgBox "Default User!", vbOKOnly + vbExclamation, "Warning..."
    List1.Text = ""
    cmdDelete.Enabled = False
    cmdchange.Enabled = False
    Exit Sub
End If
If MsgBox("Delete Username and Password?", vbYesNo + vbQuestion, "DELETING USER") = vbNo Then
    List1.Text = ""
    cmdDelete.Enabled = False
    cmdchange.Enabled = False
    Exit Sub
End If
           
    'users
    Set userCMD = New ADODB.Command
    SQLstr = "DELETE * FROM userlist where username ='" & List1.Text & "'"
    With userCMD
        .ActiveConnection = libCON
        .CommandType = adCmdText
        .CommandText = SQLstr
        .Execute
    End With
List1.clear
Call listrefresh

cmdDelete.Enabled = False
cmdchange.Enabled = False

MsgBox "Username and Password Deleted!", vbOKOnly + vbInformation, "Information"

End Sub

Private Sub Form_Load()
    dbconnect
    Call listrefresh
End Sub

Private Sub listrefresh()
'users
Set userRS = New ADODB.Recordset
userRS.Open "userlist", libCON, adOpenKeyset, adLockReadOnly

While userRS.EOF <> True
    List1.AddItem userRS!Username
    userRS.MoveNext
Wend

End Sub

Private Sub List1_Click()
If List1.Text <> Empty Then
    cmdDelete.Enabled = True
    cmdchange.Enabled = True
End If
End Sub
