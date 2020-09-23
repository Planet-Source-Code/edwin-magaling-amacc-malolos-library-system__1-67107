VERSION 5.00
Begin VB.Form frmaddbook 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&CANCEL"
      Height          =   420
      Left            =   5490
      TabIndex        =   7
      Top             =   2700
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&ADD"
      Height          =   420
      Left            =   4410
      TabIndex        =   6
      Top             =   2700
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&VIEW CATALOG"
      Height          =   420
      Left            =   495
      TabIndex        =   5
      Top             =   2700
      Width           =   1950
   End
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
      Height          =   2220
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   6990
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
         Left            =   2250
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "All"
         Top             =   405
         Width           =   2895
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
         TabIndex        =   2
         Top             =   945
         Width           =   1635
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
         Left            =   2250
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1485
         Width           =   4155
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
         Left            =   990
         TabIndex        =   8
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label6 
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
         TabIndex        =   4
         Top             =   1035
         Width           =   1635
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
         Left            =   1485
         TabIndex        =   3
         Top             =   1575
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmaddbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetListViewTo(ByVal xrs As ADODB.Recordset, Optional ByVal strSMIcons As String = "", Optional ByVal strLRGIcons As String = "", Optional ByVal clmWidth)
'ListView1.ColumnHeaders.clear
'frmborrowing.ListView1.ListItems.clear

'If xrs.EOF And xrs.BOF Then
'    Exit Sub
'End If

'xrs.MoveFirst
'Columns
xrs.MoveFirst
'While Not xrs.EOF
    'On Error Resume Next
    Set Item = frmborrowing.ListView1.ListItems.Add(, "_" & xrs.Fields(0).Value, xrs.Fields(0).Value)
    'Item.SubItems(1) = xrs!perspec_num
    Item.SubItems(1) = xrs!Title
    'Item.SubItems(2) = xrs!price
    'Item.SubItems(3) = xrs!quantity
    'Item.SubItems(4) = xrs!estimated_price
'    xrs.MoveNext
'Wend
End Sub

Public Sub SetSectionView() '(ByVal key As Integer)
    
    Set bookRS = New ADODB.Recordset
    SQLstr = "SELECT * FROM book_catalog WHERE access_no='" & Combo1.Text & "'"
    'bookRS.CursorLocation = adUseClient
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    
    SetListViewTo bookRS, 2, 2, clmWidth
End Sub


Private Sub Combo3_Click()
Combo1.clear
Combo2.clear
If Combo3.Text = "All" Then
    'accession number & title
    Set bookRS = New ADODB.Recordset
    bookRS.Open "book_catalog", libCON, adOpenKeyset, adLockReadOnly
    While bookRS.EOF <> True
        Combo1.AddItem bookRS!access_no
        Combo2.AddItem bookRS!Title
        bookRS.MoveNext
    Wend
Else
    'accession number & title
    Set bookRS = New ADODB.Recordset
    SQLstr = "select * from book_catalog where category='" & Combo3.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    While bookRS.EOF <> True
        Combo1.AddItem bookRS!access_no
        Combo2.AddItem bookRS!Title
        bookRS.MoveNext
    Wend
End If
End Sub

Private Sub Command2_Click()
SetSectionView
frmborrowing.Command2.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   'category
    Set catRS = New ADODB.Recordset
    catRS.Open "category", libCON, adOpenKeyset, adLockReadOnly
    Combo3.AddItem "All"
    While catRS.EOF <> True
        Combo3.AddItem catRS!category
        catRS.MoveNext
    Wend
   'accession number & title
    Set bookRS = New ADODB.Recordset
    bookRS.Open "book_catalog", libCON, adOpenKeyset, adLockReadOnly
    While bookRS.EOF <> True
        Combo1.AddItem bookRS!access_no
        Combo2.AddItem bookRS!Title
        bookRS.MoveNext
    Wend
    
    Command1.Enabled = False
    Command2.Enabled = False
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
    
    If Combo1.Text <> "" And Combo2.Text <> "" Then
        Command1.Enabled = True
        Command2.Enabled = True
    Else
        Command1.Enabled = False
        Command2.Enabled = False
    End If
End Sub

Private Sub Combo1_Click()
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where access_no='" & Combo1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    Combo2.Text = bookRS!Title
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
    Combo2.Text = bookRS!Title
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

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where title='" & Combo2.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    Combo1.Text = bookRS!access_no
End If
End Sub
Private Sub combo2_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    If Not auto And Combo2.Text <> "" Then
        iStart = Combo2.SelStart
        strPart = Left$(Combo2.Text, iStart)
        For iLoop = 0 To Combo2.ListCount - 1
            strItem = UCase$(Combo2.List(iLoop))
            If strItem Like UCase$(strPart & "*") And strItem <> UCase$(Combo2.Text) Then
                auto = True
                Combo2.SelText = Mid$(Combo2.List(iLoop), iStart + 1)
                Combo2.SelStart = iStart
                Combo2.SelLength = Len(Combo2.Text) - iStart
                auto = False
                Exit For
            End If
        Next iLoop
    End If
    
    If Combo1.Text <> "" And Combo2.Text <> "" Then
        Command1.Enabled = True
        Command2.Enabled = True
    Else
        Command1.Enabled = False
        Command2.Enabled = False
    End If
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        auto = True
        Combo2.SelText = ""
        auto = False
    ElseIf KeyCode = vbKeyReturn Then
        combo2_LostFocus
        Combo2.SelStart = Len(Combo2.Text)
    End If
End Sub

Private Sub combo2_LostFocus()
Dim iLoop As Integer
    If Combo2.Text <> "" Then
        For iLoop = 0 To Combo2.ListCount - 1
            If UCase$(Combo2.List(iLoop)) = UCase$(Combo2.Text) Then
                auto = True
                Combo2.Text = Combo2.List(iLoop)
                auto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub Combo2_Click()
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where title='" & Combo2.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    Combo1.Text = bookRS!access_no
End Sub

