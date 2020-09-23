VERSION 5.00
Begin VB.Form frmbookcatalog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Catalog"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   Icon            =   "frmbookcatalog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "CL&OSE"
      Height          =   465
      Left            =   6840
      TabIndex        =   14
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&FIND"
      Height          =   465
      Left            =   5760
      TabIndex        =   13
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&DELETE"
      Height          =   465
      Left            =   3600
      TabIndex        =   12
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&EDIT"
      Height          =   465
      Left            =   2520
      TabIndex        =   11
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SAVE"
      Height          =   465
      Left            =   1440
      TabIndex        =   10
      Top             =   5310
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&NEW"
      Height          =   465
      Left            =   360
      TabIndex        =   9
      Top             =   5310
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Book Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   225
      TabIndex        =   15
      Top             =   360
      Width           =   7755
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000009&
         Height          =   465
         Left            =   3510
         Picture         =   "frmbookcatalog.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   510
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frmbookcatalog.frx":0884
         Left            =   2205
         List            =   "frmbookcatalog.frx":0886
         TabIndex        =   7
         Top             =   3285
         Width           =   915
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Left            =   2205
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1755
         Width           =   3435
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
         Height          =   375
         Left            =   2205
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   2745
         Width           =   5235
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2205
         TabIndex        =   8
         Top             =   3780
         Width           =   915
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
         Height          =   375
         Left            =   2205
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   2250
         Width           =   5235
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
         Height          =   375
         Left            =   2205
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1260
         Width           =   5235
      End
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
         Left            =   2205
         TabIndex        =   2
         Top             =   765
         Width           =   3345
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Left            =   2205
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Edition :"
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
         Left            =   1080
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "No. of Copies :"
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
         TabIndex        =   22
         Top             =   3780
         Width           =   1545
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Year Published :"
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
         Left            =   225
         TabIndex        =   21
         Top             =   3330
         Width           =   1725
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Publisher :"
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
         TabIndex        =   20
         Top             =   2835
         Width           =   1110
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Author/s :"
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
         TabIndex        =   19
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Title :"
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
         TabIndex        =   18
         Top             =   1305
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Accession No. :"
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
         Left            =   315
         TabIndex        =   17
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Category :"
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
         Left            =   855
         TabIndex        =   16
         Top             =   810
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmbookcatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y As Integer

Private Sub clear()
    Combo1.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
End Sub

Private Sub disable()
    Combo1.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Combo2.Enabled = False
    Combo3.Enabled = False
End Sub

Private Sub enable()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
End Sub

Private Sub Command1_Click()
    enable
    Text1.SetFocus
    Command1.Enabled = False
    Command2.Enabled = True
    Command5.Enabled = False
    Command6.Caption = "&CANCEL"
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text1.SetFocus
    Exit Sub
End If
If Combo1.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Combo1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text2.SetFocus
    Exit Sub
End If
If Text5.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text5.SetFocus
    Exit Sub
End If
If Text3.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text3.SetFocus
    Exit Sub
End If
If Text4.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Text4.SetFocus
    Exit Sub
End If
If Combo2.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Combo2.SetFocus
    Exit Sub
End If
If Combo3.Text = "" Then
    MsgBox "Complete neccessary information", vbExclamation
    Combo3.SetFocus
    Exit Sub
End If

If Command2.Caption = "&SAVE" Then
    Set bookRS = New ADODB.Recordset
    SQLstr = "select access_no from book_catalog where access_no='" & Text1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If Not bookRS.EOF And Not bookRS.BOF Then
        MsgBox "Accession No. already exist!", vbExclamation
        Text1.SetFocus
    Exit Sub
End If
    If MsgBox("Save Book Catalog?", vbYesNo + vbQuestion) = vbYes Then
        Set bookRS = New ADODB.Recordset
        bookRS.Open "book_catalog", libCON, adOpenKeyset, adLockOptimistic
        With bookRS
            .AddNew
            !access_no = Text1.Text
            !category = Combo1.Text
            !Title = Text2.Text
            !Edition = Text5.Text
            !Author = Text3.Text
            !Publisher = Text4.Text
            !yr_publish = Combo3.Text
            !no_copy = Combo2.Text
            !available_copy = Combo2.Text
            .Update
            .Close
        End With
        MsgBox "Book Catalog Successfully Saved!", vbInformation
    End If
Else
    If MsgBox("Update Book Catalog?", vbYesNo + vbQuestion) = vbYes Then
        Set bookRS = New ADODB.Recordset
        SQLstr = "Select * from book_catalog where access_no='" & Text1.Text & "'"
        bookRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
        With bookRS
            !category = Combo1.Text
            !Title = Text2.Text
            !Edition = Text5.Text
            !Author = Text3.Text
            !Publisher = Text4.Text
            !yr_publish = Combo3.Text
            !no_copy = Combo2.Text
            !available_copy = Combo2.Text
            .Update
            .Close
        End With
        MsgBox "Book Catalog Successfully Updated!", vbInformation
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
    Text1.Enabled = False
    Combo1.SetFocus
    Command2.Enabled = True
    Command2.Caption = "&UPDATE"
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Visible = False
End Sub

Private Sub Command4_Click()
    If MsgBox("Sure To Delete Book Catalog?", vbQuestion + vbYesNo) = vbYes Then
        Set bookCMD = New ADODB.Command
        SQLstr = "Delete * from book_catalog where access_no='" & Text1.Text & "'"
        With bookCMD
            .ActiveConnection = libCON
            .CommandType = adCmdText
            .CommandText = SQLstr
            .Execute
        End With
        clear
        
        MsgBox "Book Catalog Successfully Deleted!", vbInformation
        
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
    Text1.Enabled = True
    Text1.SetFocus
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
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where access_no='" & Text1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If bookRS.EOF And bookRS.BOF Then
        MsgBox "Accession Number Not Found!", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    With bookRS
        Combo1.Text = !category
        Text2.Text = !Title
        Text3.Text = !Author
        Text4.Text = !Publisher
        Text5.Text = !Edition
        Combo2.Text = !no_copy
        Combo3.Text = !yr_publish
    End With
        Command3.Enabled = True
        Command4.Enabled = True
End Sub

Private Sub Form_Load()
    dbconnect
    clear
    disable
    
    Set catRS = New ADODB.Recordset
    catRS.Open "category", libCON, adOpenKeyset, adLockReadOnly
    While catRS.EOF <> True
        Combo1.AddItem catRS!category
        catRS.MoveNext
    Wend
    
    x = 1
    While x <= 10
        Combo2.AddItem x
        x = x + 1
    Wend
    
    y = 1601
    While y <= 9999
        Combo3.AddItem y
        y = y + 1
    Wend
    
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command7.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Command7.Visible = True Then
    If KeyAscii = 13 Then
        Command7.Value = True
    End If
End If
End Sub
