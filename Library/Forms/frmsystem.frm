VERSION 5.00
Begin VB.Form frmsystem 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changing Overdue Fee"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3075
   Icon            =   "frmsystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3075
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&APPLY"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CL&OSE"
      Height          =   375
      Left            =   1710
      TabIndex        =   2
      Top             =   1350
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   675
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   630
      Width           =   1770
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Overdue Fee Per Day"
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
      Height          =   240
      Left            =   405
      TabIndex        =   0
      Top             =   225
      Width           =   2265
   End
End
Attribute VB_Name = "frmsystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Set feeCMD = New ADODB.Command
    SQLstr = "delete * from charge"
    With feeCMD
        .ActiveConnection = libCON
        .CommandType = adCmdText
        .CommandText = SQLstr
        .Execute
    End With
    
    Set feeRS = New ADODB.Recordset
    feeRS.Open "charge", libCON, adOpenKeyset, adLockOptimistic
    With feeRS
        .AddNew
        !charge_fee = Text1.Text
        .Update
        .Close
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dbconnect
    
    Set feeRS = New ADODB.Recordset
    feeRS.Open "charge", libCON, adOpenKeyset, adLockReadOnly
    Text1.Text = feeRS!charge_fee
    Text1.Text = Format(Text1.Text, "###.00")
End Sub

Private Sub Text1_Change()
    If Text1.Text <> "" Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub
