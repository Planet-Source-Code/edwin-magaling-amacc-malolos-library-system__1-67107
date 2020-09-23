VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfind 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finding Book"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10500
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Find Book By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7830
      TabIndex        =   3
      Top             =   180
      Width           =   2085
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Author"
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
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   810
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Title"
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
         Left            =   360
         TabIndex        =   4
         Top             =   405
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CL&OSE"
      Height          =   510
      Left            =   8865
      TabIndex        =   2
      Top             =   4365
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   465
      Left            =   7110
      Picture         =   "frmfind.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   675
      Width           =   600
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
      Height          =   465
      Left            =   2205
      TabIndex        =   0
      Top             =   675
      Width           =   4830
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2130
      Left            =   225
      TabIndex        =   7
      Top             =   2070
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   3757
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accession No."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Edition"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Publisher"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Year Published"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "No. of Copy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "On Shelves"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "On Hand"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "List of Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1620
      Width           =   1770
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   180
      TabIndex        =   6
      Top             =   810
      Width           =   75
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetListViewTo(ByVal xrs As ADODB.Recordset, Optional ByVal strSMIcons As String = "", Optional ByVal strLRGIcons As String = "", Optional ByVal clmWidth)
ListView1.ListItems.clear
xrs.MoveFirst
While Not xrs.EOF
    Set Item = ListView1.ListItems.Add(, "_" & xrs.Fields(0).Value, xrs.Fields(0).Value)
    'Item.SubItems(1) = xrs!access_no
    Item.SubItems(1) = xrs!Title
    Item.SubItems(2) = xrs!Edition
    Item.SubItems(3) = xrs!Author
    Item.SubItems(4) = xrs!Publisher
    Item.SubItems(5) = xrs!yr_publish
    Item.SubItems(6) = xrs!no_copy
    Item.SubItems(7) = xrs!available_copy
    Item.SubItems(8) = xrs!borrow_copy
    xrs.MoveNext
Wend
End Sub

Public Sub SetSectionViewTitle()
    
    Set bookRS = New ADODB.Recordset
    SQLstr = "SELECT * FROM book_catalog WHERE title='" & Text1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If bookRS.EOF And bookRS.BOF Then
        MsgBox "No title found!", vbInformation
        Exit Sub
    End If
    SetListViewTo bookRS, 2, 2, clmWidth
    
End Sub

Public Sub SetSectionViewAuthor()
    
    Set bookRS = New ADODB.Recordset
    SQLstr = "SELECT * FROM book_catalog WHERE author='" & Text1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If bookRS.EOF And bookRS.BOF Then
        MsgBox "No Author found!", vbInformation
        Exit Sub
    End If
    
    SetListViewTo bookRS, 2, 2, clmWidth
End Sub



Private Sub Command1_Click()
    If Option1.Value = True Then
        SetSectionViewTitle
    Else
        SetSectionViewAuthor
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Option1.Value = True
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        Label1.Caption = "Search by Title :"
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Label1.Caption = "Search by Author :"
    End If
End Sub

Private Sub Text1_Change()
ListView1.ListItems.clear
End Sub
