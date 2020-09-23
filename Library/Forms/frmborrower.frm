VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmborrower 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrower's Record"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmborrower.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Borrower's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   225
      TabIndex        =   16
      Top             =   180
      Width           =   6630
      Begin VB.ComboBox Combo1 
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
         Left            =   1935
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2115
         Width           =   1275
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Faculty / Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3330
         TabIndex        =   6
         Top             =   1620
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1980
         TabIndex        =   5
         Top             =   1620
         Width           =   1140
      End
      Begin VB.TextBox Text5 
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
         Height          =   360
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "Text5"
         Top             =   405
         Width           =   1590
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000009&
         Height          =   465
         Left            =   3600
         Picture         =   "frmborrower.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   510
      End
      Begin VB.TextBox Text1 
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
         Left            =   1935
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   900
         Width           =   1905
      End
      Begin VB.TextBox Text2 
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
         Left            =   3915
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   5805
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   900
         Width           =   600
      End
      Begin VB.TextBox Text4 
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
         Left            =   1935
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   2655
         Width           =   4470
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Left            =   1935
         TabIndex        =   9
         Top             =   3240
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   " Course :"
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
         Left            =   810
         TabIndex        =   25
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Borrower ID :"
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
         TabIndex        =   24
         Top             =   495
         Width           =   1350
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Name :"
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
         Left            =   990
         TabIndex        =   23
         Top             =   945
         Width           =   750
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Status :"
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
         Left            =   945
         TabIndex        =   22
         Top             =   1665
         Width           =   780
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Address :"
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
         Left            =   720
         TabIndex        =   21
         Top             =   2745
         Width           =   1005
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Contact No. :"
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
         Left            =   450
         TabIndex        =   20
         Top             =   3330
         Width           =   1365
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Lastname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2430
         TabIndex        =   19
         Top             =   1260
         Width           =   885
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Firstname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4320
         TabIndex        =   18
         Top             =   1260
         Width           =   885
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "MI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5985
         TabIndex        =   17
         Top             =   1260
         Width           =   210
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CL&OSE"
      Height          =   465
      Left            =   5760
      TabIndex        =   15
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&FIND"
      Height          =   465
      Left            =   4815
      TabIndex        =   14
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&DELETE"
      Height          =   465
      Left            =   3285
      TabIndex        =   13
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&EDIT"
      Height          =   465
      Left            =   2340
      TabIndex        =   12
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SAVE"
      Height          =   465
      Left            =   1395
      TabIndex        =   11
      Top             =   4410
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&NEW"
      Height          =   465
      Left            =   450
      TabIndex        =   10
      Top             =   4410
      Width           =   915
   End
End
Attribute VB_Name = "frmborrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear()
    Text5.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Option1.Value = False
    Option2.Value = False
    Combo1.Text = ""
    Text4.Text = ""
    MaskEdBox2.Text = "(   )    -    "
End Sub

Private Sub disable()
    Text5.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Text4.Enabled = False
    Combo1.Enabled = False
    MaskEdBox2.Enabled = False
End Sub

Private Sub enable()
    Text5.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    Text4.Enabled = True
    'Combo1.Enabled = True
    MaskEdBox2.Enabled = True
End Sub

Private Sub Command1_Click()
    enable
    Text5.SetFocus
    Command1.Enabled = False
    Command2.Enabled = True
    Command5.Enabled = False
    Command6.Caption = "&CANCEL"
End Sub

Private Sub Command2_Click()
If Text5.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text5.SetFocus
    Exit Sub
End If
If Text1.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text2.SetFocus
    Exit Sub
End If
If Text3.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text3.SetFocus
    Exit Sub
End If
If Option1.Value = False And Option2.Value = False Then
    MsgBox "Complete neccessary information", vbExclamation
    Option1.SetFocus
    Exit Sub
End If
If Option1.Value = True And Combo1.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Combo1.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text4.SetFocus
    Exit Sub
End If
      
If Command2.Caption = "&SAVE" Then
    Set borrowerRS = New ADODB.Recordset
    SQLstr = "select borrower_id from borrower_record where borrower_id='" & Text5.Text & "'"
    borrowerRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If Not borrowerRS.EOF And Not borrowerRS.BOF Then
        MsgBox "Borrower ID already exist!", vbExclamation
        Text5.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Save Borrower Record?", vbYesNo + vbQuestion) = vbYes Then
        Set borrowerRS = New ADODB.Recordset
        borrowerRS.Open "borrower_record", libCON, adOpenKeyset, adLockOptimistic
        With borrowerRS
            .AddNew
            !borrower_id = Text5.Text
            !lname = Text1.Text
            !fname = Text2.Text
            !mI = Text3.Text
            If Option1.Value = True Then
                !Status = "Student"
            Else
                !Status = "Faculty / Employee"
            End If
            !course = Combo1.Text
            !Add = Text4.Text
            !contact = MaskEdBox2.Text
            .Update
            .Close
        End With
        MsgBox "Borrower Record Successfully Saved!", vbInformation
    End If
Else
    If MsgBox("Update Borrower Record?", vbYesNo + vbQuestion) = vbYes Then
        Set borrowerRS = New ADODB.Recordset
        SQLstr = "Select * from borrower_record where borrower_id='" & Text5.Text & "'"
        borrowerRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
        With borrowerRS
            !lname = Text1.Text
            !fname = Text2.Text
            !mI = Text3.Text
            If Option1.Value = True Then
                !Status = "Student"
            Else
                !Status = "Faculty / Employee"
            End If
            !course = Combo1.Text
            !Add = Text4.Text
            !contact = MaskEdBox2.Text
            .Update
            .Close
        End With
        MsgBox "Borrower Record Successfully Updated!", vbInformation
    End If
End If

        clear
        disable
        Command1.Enabled = True
        Command2.Enabled = False
        Command2.Caption = "&SAVE"
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = True
        Command6.Caption = "CL&OSE"
End Sub

Private Sub Command3_Click()
    enable
    Text5.Enabled = False
    Text1.SetFocus
    Combo1.Enabled = True
    Command2.Enabled = True
    Command2.Caption = "&UPDATE"
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Visible = False
End Sub

Private Sub Command4_Click()
    If MsgBox("Sure To Delete Borrower Record?", vbQuestion + vbYesNo) = vbYes Then
        Set borrowerCMD = New ADODB.Command
        SQLstr = "Delete * from borrower_record where borrower_id='" & Text5.Text & "'"
        With borrowerCMD
            .ActiveConnection = libCON
            .CommandType = adCmdText
            .CommandText = SQLstr
            .Execute
        End With
        clear
        
        MsgBox "Borrower Record Successfully Deleted!", vbInformation
        
        Command1.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = True
        Command6.Caption = "CL&OSE"
        Command7.Visible = False
    End If
End Sub

Private Sub Command5_Click()
    Text5.Enabled = True
    Text5.SetFocus
    Command1.Enabled = False
    Command5.Enabled = False
    Command6.Caption = "&CANCEL"
    Command7.Visible = True
End Sub

Private Sub Command6_Click()
    If Command6.Caption = "CL&OSE" Then
        Unload Me
    Else
        clear
        disable
        Command1.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = True
        Command6.Caption = "CL&OSE"
        Command7.Visible = False
    End If
End Sub

Private Sub Command7_Click()
    Set borrowerRS = New ADODB.Recordset
    SQLstr = "Select * from borrower_record where borrower_id='" & Text5.Text & "'"
    borrowerRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If borrowerRS.EOF And borrowerRS.BOF Then
        MsgBox "Borrower ID not valid!", vbExclamation
        Text5.SetFocus
        Exit Sub
    End If
    With borrowerRS
        Text1.Text = !lname
        Text2.Text = !fname
        Text3.Text = !mI
        If !Status = "Student" Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
        Combo1.Text = !course
        Text4.Text = !Add
        MaskEdBox2.Text = !contact
    End With
        Command3.Enabled = True
        Command4.Enabled = True
End Sub

Private Sub Form_Load()
    
    dbconnect
    clear
    disable
    
    Set courseRS = New ADODB.Recordset
    courseRS.Open "course", libCON, adOpenKeyset, adLockReadOnly
    While courseRS.EOF <> True
        Combo1.AddItem courseRS!course
        courseRS.MoveNext
    Wend
        
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Visible = False
End Sub

Private Sub Option1_Click()
If Option1.Enabled = True Then
    Combo1.Enabled = True
    Combo1.Text = ""
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled = True Then
    Combo1.Enabled = False
    Combo1.Text = "n/a"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Command7.Visible = True Then
    If KeyAscii = 13 Then
        Command7.Value = True
    End If
End If
End Sub
