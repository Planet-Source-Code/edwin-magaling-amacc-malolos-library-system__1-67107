VERSION 5.00
Begin VB.Form frmaddcategory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding New Book Category"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "frmaddcategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   1305
      TabIndex        =   2
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CL&OSE"
      Height          =   375
      Left            =   2655
      TabIndex        =   1
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
      Left            =   450
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   630
      Width           =   3930
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "New Book Category"
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
      Left            =   1350
      TabIndex        =   3
      Top             =   225
      Width           =   2070
   End
End
Attribute VB_Name = "frmaddcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Set catRS = New ADODB.Recordset
    SQLstr = "select * from category where category='" & Text1.Text & "'"
    catRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If Not catRS.EOF And Not catRS.BOF Then
        MsgBox "Book category already exist!", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    
    Set catRS = New ADODB.Recordset
    catRS.Open "category", libCON, adOpenKeyset, adLockOptimistic
    With catRS
        .AddNew
        !category = Text1.Text
        .Update
        .Close
    End With
    
    frmcategory.List1.clear
    Set catRS = New ADODB.Recordset
    catRS.Open "category", libCON, adOpenKeyset, adLockReadOnly
    While catRS.EOF <> True
        frmcategory.List1.AddItem catRS!category
        catRS.MoveNext
    Wend
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dbconnect
    Text1.Text = ""
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

