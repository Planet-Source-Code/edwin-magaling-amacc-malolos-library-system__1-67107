VERSION 5.00
Begin VB.Form frmaddcourse 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding New Course"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   Icon            =   "frmaddcourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
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
      Left            =   585
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   630
      Width           =   1950
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CL&OSE"
      Height          =   375
      Left            =   1710
      TabIndex        =   2
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "New Course"
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
      Left            =   900
      TabIndex        =   3
      Top             =   225
      Width           =   1260
   End
End
Attribute VB_Name = "frmaddcourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Set courseRS = New ADODB.Recordset
    SQLstr = "select * from course where course='" & Text1.Text & "'"
    courseRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If Not courseRS.EOF And Not courseRS.BOF Then
        MsgBox "Student Course already exist!", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    
    Set courseRS = New ADODB.Recordset
    courseRS.Open "course", libCON, adOpenKeyset, adLockOptimistic
    With courseRS
        .AddNew
        !course = Text1.Text
        .Update
        .Close
    End With
    
    frmcourse.List1.clear
    Set courseRS = New ADODB.Recordset
    courseRS.Open "course", libCON, adOpenKeyset, adLockReadOnly
    While courseRS.EOF <> True
        frmcourse.List1.AddItem courseRS!course
        courseRS.MoveNext
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
