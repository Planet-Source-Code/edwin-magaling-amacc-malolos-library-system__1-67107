VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmborrowing 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrowing of Books"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "frmborrowing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Book's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   180
      TabIndex        =   14
      Top             =   3825
      Width           =   10050
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2250
         TabIndex        =   20
         Top             =   990
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19464193
         CurrentDate     =   38309
      End
      Begin VB.TextBox Text4 
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
         Left            =   5085
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   405
         Width           =   4245
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
         Left            =   2250
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Due Date :"
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
         TabIndex        =   19
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label9 
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
         Left            =   4320
         TabIndex        =   17
         Top             =   495
         Width           =   600
      End
      Begin VB.Label Label7 
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
         Left            =   450
         TabIndex        =   16
         Top             =   495
         Width           =   1635
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&BORROW"
      Height          =   465
      Left            =   6705
      TabIndex        =   13
      Top             =   5805
      Width           =   1005
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   180
      TabIndex        =   10
      Top             =   2205
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   2646
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Accession No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Edition"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Publisher"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year Published"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SAVE"
      Height          =   465
      Left            =   7740
      TabIndex        =   9
      Top             =   5805
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CL&OSE"
      Height          =   465
      Left            =   8775
      TabIndex        =   8
      Top             =   5805
      Width           =   1050
   End
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
      Height          =   1545
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   10050
      Begin VB.TextBox Text3 
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
         Left            =   7470
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   900
         Width           =   2265
      End
      Begin VB.TextBox Text2 
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
         Left            =   1980
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   900
         Width           =   4245
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
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1545
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000009&
         Height          =   465
         Left            =   3600
         Picture         =   "frmborrowing.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Date :"
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
         Left            =   6660
         TabIndex        =   7
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Left            =   7425
         TabIndex        =   6
         Top             =   450
         Width           =   75
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
         Left            =   1080
         TabIndex        =   5
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
         Left            =   6525
         TabIndex        =   4
         Top             =   990
         Width           =   780
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
         Left            =   495
         TabIndex        =   3
         Top             =   405
         Width           =   1350
      End
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Borrowed Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3735
      TabIndex        =   21
      Top             =   1800
      Width           =   2310
   End
End
Attribute VB_Name = "frmborrowing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xcat, xedition, xauthor, xpublisher, xyr_publish As String
Dim xavail_copy As Integer

Private Sub SetListViewTo(ByVal xrs As ADODB.Recordset, Optional ByVal strSMIcons As String = "", Optional ByVal strLRGIcons As String = "", Optional ByVal clmWidth)
ListView1.ListItems.clear
xrs.MoveFirst
While Not xrs.EOF
    Set Item = ListView1.ListItems.Add(, "_" & xrs.Fields(0).Value, xrs.Fields(0).Value)
    Item.SubItems(1) = xrs!access_no
    Item.SubItems(2) = xrs!Title
    Item.SubItems(3) = xrs!Edition
    Item.SubItems(4) = xrs!Author
    Item.SubItems(5) = xrs!Publisher
    Item.SubItems(6) = xrs!yr_publish
    xrs.MoveNext
Wend
End Sub

Public Sub SetSectionView()
    
    Set currentRS = New ADODB.Recordset
    SQLstr = "SELECT * FROM current_borrow WHERE borrower_id='" & Text1.Text & "'"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If currentRS.EOF And currentRS.BOF Then
        Exit Sub
    End If
    SetListViewTo currentRS, 2, 2, clmWidth
End Sub

Private Sub clear()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Combo1.Text = ""
    DTPicker1.Value = Date
End Sub

Private Sub disable()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Combo1.Enabled = False
    DTPicker1.Enabled = False
End Sub

Private Sub Combo1_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not auto And Combo1.Text <> "" Then
        iStart = Combo1.SelStart
        strPart = Left$(Combo1.Text, iStart)
        For iLoop = 0 To Combo1.ListCount - 1
            strItem = UCase$(Combo1.List(iLoop))
            If strItem Like UCase$(strPart & "*") And strItem <> UCase$(Combo1.Text) Then
                auto = True
                Combo1.SelText = Mid$(Combo1.List(iLoop), iStart + 1)
                Combo1.SelStart = iStart
                Combo1.SelLength = Len(Combo1.Text) - iStart
                auto = False
                Exit For
            End If
        Next iLoop
    End If

If Combo1.Text = "" Then
    Text4.Text = ""
    xcat = ""
    xedition = ""
    xauthor = ""
    xpublisher = ""
    xyr_publish = ""
End If
If Combo1.Text <> "" And Text4.Text <> "" Then
    Command2.Enabled = True
Else
    Command2.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where access_no='" & Combo1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    With bookRS
        Text4.Text = !Title
        xcat = !category
        xedition = !Edition
        xauthor = !Author
        xpublisher = !Publisher
        xyr_publish = !yr_publish
        xavail_copy = !available_copy
    End With
    
If Combo1.Text <> "" And Text4.Text <> "" Then
    Command2.Enabled = True
Else
    Command2.Enabled = False
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        auto = True
        Combo1.SelText = ""
        auto = False
    ElseIf KeyCode = vbKeyReturn Then
        Combo1_LostFocus
        Combo1.SelStart = Len(Combo1.Text)
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where access_no='" & Combo1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If bookRS.EOF And bookRS.BOF Then
        Exit Sub
    End If
    With bookRS
        Text4.Text = !Title
        xcat = !category
        xedition = !Edition
        xauthor = !Author
        xpublisher = !Publisher
        xyr_publish = !yr_publish
        xavail_copy = !available_copy
    End With
    
    If Combo1.Text <> "" And Text4.Text <> "" Then
    Command2.Enabled = True
    Else
    Command2.Enabled = False
    End If
End If
End Sub

Private Sub Combo1_LostFocus()
Dim iLoop As Integer
    If Combo1.Text <> "" Then
        For iLoop = 0 To Combo1.ListCount - 1
            If UCase$(Combo1.List(iLoop)) = UCase$(Combo1.Text) Then
                auto = True
                Combo1.Text = Combo1.List(iLoop)
                auto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Command1_Click()
    Text1.Enabled = True
    Text1.SetFocus
    Command1.Enabled = False
    Command6.Caption = "&CANCEL"
    Command7.Visible = True
End Sub

Private Sub Command3_Click()
    Text1.Enabled = False
    Command7.Visible = False
    frmaddbook.Show vbModal
End Sub

Private Sub Command2_Click()
    If xavail_copy = 0 Then
        MsgBox "There is no available copy of this book!", vbInformation
        Combo1.SetFocus
        Exit Sub
    End If
           
    Set currentRS = New ADODB.Recordset
    SQLstr = "Select * from current_borrow where access_no='" & Combo1.Text & "'" & " and borrower_id='" & Text1.Text & "'"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If Not currentRS.EOF And Not currentRS.BOF Then
        MsgBox "Borrower's cannot borrow 2 same book title!", vbInformation
        Combo1.SetFocus
        Exit Sub
    End If
    
    Set currentRS = New ADODB.Recordset
    currentRS.Open "current_borrow", libCON, adOpenKeyset, adLockOptimistic
    With currentRS
        .AddNew
        !access_no = Combo1.Text
        !Title = Text4.Text
        !category = xcat
        !Edition = xedition
        !Author = xauthor
        !Publisher = xpublisher
        !yr_publish = xyr_publish
        !borrow_date = Label4.Caption
        If DTPicker1.Enabled = True Then
            !due_date = DTPicker1.Value
        End If
        !borrower_id = Text1.Text
        !Name = Text2.Text
        !Status = Text3.Text
        .Update
        .Close
    End With
        
    Set bookRS = New ADODB.Recordset
    SQLstr = "select * from book_catalog where access_no='" & Combo1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
    With bookRS
        !available_copy = xavail_copy - 1
        !borrow_copy = !borrow_copy + 1
        .Update
        .Close
    End With

        SetSectionView
        MsgBox "Library Transaction Successfully Saved!", vbInformation
        
        If ListView1.ListItems.Count = 3 Then
            'MsgBox "Borrower reach the maximum book that can be borrowed at a time!", vbExclamation
            clear
            disable
            ListView1.ListItems.clear
            Command1.Enabled = True
            Command2.Enabled = False
            Command6.Caption = "CL&OSE"
            Exit Sub
        End If
        
        If MsgBox("Borrow Another Book?", vbQuestion + vbYesNo) = vbYes Then
            Combo1.Text = ""
            Combo1.SetFocus
            Exit Sub
        End If
            clear
            disable
            ListView1.ListItems.clear
            Command1.Enabled = True
            Command2.Enabled = False
            Command6.Caption = "CL&OSE"
End Sub

Private Sub Command6_Click()
If Command6.Caption = "CL&OSE" Then
    Unload Me
Else
    clear
    disable
    Command1.Enabled = True
    Command2.Enabled = False
    Command6.Caption = "CL&OSE"
    Command7.Visible = False
End If
End Sub

Private Sub Command7_Click()
    Set borrowerRS = New ADODB.Recordset
    SQLstr = "Select * from borrower_record where borrower_id='" & Text1.Text & "'"
    borrowerRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If borrowerRS.EOF And borrowerRS.BOF Then
        MsgBox "Borrower ID not valid!", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    With borrowerRS
        Text2.Text = !lname & ", " & !fname & " " & !mI
        Text3.Text = !Status
    End With
    
    SetSectionView
    If ListView1.ListItems.Count = 3 Then
        MsgBox "Borrower Already Borrowed 3 Books!", vbExclamation
        clear
        ListView1.ListItems.clear
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
    End If
    If Text3.Text = "Faculty / Employee" Then
        DTPicker1.Enabled = False
    Else
        DTPicker1.Enabled = True
    End If
    Combo1.Enabled = True
    Combo1.SetFocus
    Text1.Enabled = False
    Command7.Visible = False
End Sub

Private Sub Form_Load()
    dbconnect
    clear
    disable
    Label4.Caption = Date
    Command2.Enabled = False
    Command7.Visible = False

   'accession number
    Set bookRS = New ADODB.Recordset
    bookRS.Open "book_catalog", libCON, adOpenKeyset, adLockReadOnly
    While bookRS.EOF <> True
        Combo1.AddItem bookRS!access_no
        bookRS.MoveNext
    Wend
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command7.Value = True
End If
End Sub

