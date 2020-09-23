VERSION 5.00
Begin VB.Form frmcategory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Updating Book Category"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   Icon            =   "frmcategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
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
      Left            =   225
      TabIndex        =   3
      Top             =   450
      Width           =   3930
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CL&OSE"
      Height          =   435
      Left            =   4275
      TabIndex        =   2
      Top             =   2430
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&DELETE"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4275
      TabIndex        =   1
      Top             =   900
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
      Height          =   435
      Left            =   4275
      TabIndex        =   0
      Top             =   450
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Book Category"
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
      Left            =   990
      TabIndex        =   4
      Top             =   90
      Width           =   2355
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmaddcategory.Show vbModal
End Sub

Private Sub Command2_Click()
 Set catCMD = New ADODB.Command
    SQLstr = "DELETE * FROM category where category ='" & List1.Text & "'"
    With catCMD
        .ActiveConnection = libCON
        .CommandType = adCmdText
        .CommandText = SQLstr
        .Execute
    End With
List1.clear
listrefresh

Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dbconnect
    listrefresh
End Sub

Private Sub List1_Click()
If List1.Text <> Empty Then
    Command2.Enabled = True
End If
End Sub

Private Sub listrefresh()
Set catRS = New ADODB.Recordset
catRS.Open "category", libCON, adOpenKeyset, adLockReadOnly
While catRS.EOF <> True
    List1.AddItem catRS!category
    catRS.MoveNext
Wend
End Sub
