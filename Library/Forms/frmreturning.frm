VERSION 5.00
Begin VB.Form frmreturning 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returning of Books"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmreturning.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10455
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Caption         =   "In Case of Overdue Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   225
      TabIndex        =   18
      Top             =   4320
      Width           =   10050
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Label13"
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
         Left            =   7650
         TabIndex        =   24
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Label12"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Label11"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total Charge :"
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
         Left            =   5985
         TabIndex        =   21
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "No. of Days :"
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
         Left            =   1710
         TabIndex        =   20
         Top             =   405
         Width           =   1350
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Charge Fee Per Day :"
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
         TabIndex        =   19
         Top             =   810
         Width           =   2250
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "Borrowing Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   180
      TabIndex        =   15
      Top             =   3015
      Width           =   10050
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   6435
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   495
         Width           =   1545
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "Text4"
         Top             =   495
         Width           =   1545
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Date Borrowed :"
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
         TabIndex        =   17
         Top             =   585
         Width           =   1680
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
         Left            =   5085
         TabIndex        =   16
         Top             =   540
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CL&OSE"
      Height          =   465
      Left            =   9000
      TabIndex        =   5
      Top             =   5850
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SAVE"
      Height          =   465
      Left            =   7965
      TabIndex        =   4
      Top             =   5850
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&RETURN"
      Height          =   465
      Left            =   6930
      TabIndex        =   3
      Top             =   5850
      Width           =   1005
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
      Height          =   1500
      Left            =   180
      TabIndex        =   7
      Top             =   1485
      Width           =   10050
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2250
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   855
         Width           =   4245
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   7740
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   855
         Width           =   2085
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
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label5 
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
         Left            =   6795
         TabIndex        =   14
         Top             =   945
         Width           =   780
      End
      Begin VB.Label Label4 
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
         Left            =   1350
         TabIndex        =   13
         Top             =   900
         Width           =   750
      End
      Begin VB.Label Label3 
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
         Left            =   720
         TabIndex        =   10
         Top             =   450
         Width           =   1350
      End
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
      Height          =   1185
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   10050
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5355
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   405
         Width           =   4110
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
         ItemData        =   "frmreturning.frx":0442
         Left            =   2250
         List            =   "frmreturning.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label Label2 
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
         Left            =   4590
         TabIndex        =   8
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
         TabIndex        =   6
         Top             =   495
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmreturning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xaccess_no, xborrower_id As String

Private Sub clear()
    Combo1.Text = ""
    Text1.Text = ""
    Combo2.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = Date
    Text5.Text = Date
    Label11.Caption = "0"
    Label12.Caption = "0.00"
    Label13.Caption = "0.00"
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
        Text1.Text = ""
        Combo2.Text = ""
    End If
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Combo2.Enabled = True
    Else
        Combo2.Enabled = False
    End If
End Sub

Private Sub Combo1_Click()
    Set currentRS = New ADODB.Recordset
    SQLstr = "Select * from current_borrow where access_no='" & Combo1.Text & "'"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If currentRS.EOF And currentRS.BOF Then
        Exit Sub
    End If
    With currentRS
        Text1.Text = !Title
        Combo2.clear
        xborrower_id = ""
        While .EOF <> True
            If !borrower_id = xborrower_id Then
            .MoveNext
            Else
            Combo2.AddItem !borrower_id
            xborrower_id = !borrower_id
            .MoveNext
            End If
        Wend
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = Date
        Text5.Text = Date
    End With
    
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Combo2.Enabled = True
    Else
        Combo2.Enabled = False
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
    Set currentRS = New ADODB.Recordset
    SQLstr = "Select * from current_borrow where access_no='" & Combo1.Text & "'"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    If currentRS.EOF And currentRS.BOF Then
        Exit Sub
    End If
    With currentRS
        Text1.Text = !Title
        Combo2.clear
        xborrower_id = ""
        While .EOF <> True
            If !borrower_id = xborrower_id Then
            .MoveNext
            Else
            Combo2.AddItem !borrower_id
            xborrower_id = !borrower_id
            .MoveNext
            End If
        Wend
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = Date
        Text5.Text = Date
    End With
    
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Combo2.Enabled = True
    Else
        Combo2.Enabled = False
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

If Combo2.Text = "" Then
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = Date
    Text5.Text = Date
End If
End Sub

Private Sub Combo2_Click()
Set currentRS = New ADODB.Recordset
SQLstr = "select * from current_borrow where borrower_id='" & Combo2.Text & "'"
currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
If currentRS.EOF And currentRS.BOF Then
    Exit Sub
End If
With currentRS
    Text2.Text = !Name
    Text3.Text = !Status
    Text4.Text = !borrow_date
    On Error Resume Next
    Text5.Text = !due_date
End With
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

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set currentRS = New ADODB.Recordset
SQLstr = "select * from current_borrow where borrower_id='" & Combo2.Text & "'"
currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
If currentRS.EOF And currentRS.BOF Then
    Exit Sub
End If
With currentRS
    Text2.Text = !Name
    Text3.Text = !Status
    Text4.Text = !borrow_date
    On Error Resume Next
    Text5.Text = !due_date
End With
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

Private Sub Command1_Click()
     If currentRS.EOF And currentRS.BOF Then
        MsgBox "There is no books currently borrowed!", vbExclamation
        Exit Sub
    End If
    Combo1.Enabled = True
    Combo1.SetFocus
    Command1.Enabled = False
    Command6.Caption = "&CANCEL"
End Sub

Private Sub Command2_Click()
    Set bookRS = New ADODB.Recordset
    SQLstr = "Select * from book_catalog where access_no='" & Combo1.Text & "'"
    bookRS.Open SQLstr, libCON, adOpenKeyset, adLockOptimistic
    With bookRS
        !available_copy = !available_copy + 1
        !borrow_copy = !borrow_copy - 1
        .Update
        .Close
    End With
    
    Set currentCMD = New ADODB.Command
    SQLstr = "Delete * from current_borrow where access_no='" & Combo1.Text & "'" & " and borrower_id='" & Combo2.Text & "'"
        With currentCMD
            .ActiveConnection = libCON
            .CommandType = adCmdText
            .CommandText = SQLstr
            .Execute
        End With
    clear
    Combo1.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command6.Caption = "CL&OSE"
        
    MsgBox "Library Transaction Successfully Saved!", vbInformation
    
    Set currentRS = New ADODB.Recordset
    currentRS.Open "current_borrow", libCON, adOpenKeyset, adLockReadOnly
    Combo1.clear
    xaccess_no = ""
    While currentRS.EOF <> True
        If currentRS!access_no = xaccess_no Then
            currentRS.MoveNext
        Else
            Combo1.AddItem currentRS!access_no
            xaccess_no = currentRS!access_no
            currentRS.MoveNext
        End If
    Wend
End Sub

Private Sub Command6_Click()
    If Command6.Caption = "CL&OSE" Then
    Unload Me
Else
    clear
    Combo1.Enabled = False
    Combo2.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command6.Caption = "CL&OSE"
End If
End Sub

Private Sub Form_Load()
    dbconnect
    clear
    Combo1.Enabled = False
    Combo2.Enabled = False
    Command2.Enabled = False

    Set currentRS = New ADODB.Recordset
    currentRS.Open "current_borrow", libCON, adOpenKeyset, adLockReadOnly
    xaccess_no = ""
    While currentRS.EOF <> True
        If currentRS!access_no = xaccess_no Then
            currentRS.MoveNext
        Else
        Combo1.AddItem currentRS!access_no
        xaccess_no = currentRS!access_no
        currentRS.MoveNext
        End If
    Wend
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If

End Sub

Private Sub Text5_Change()
If Date > DateValue(Text5.Text) Then
    Label11.Caption = Date - DateValue(Text5.Text)
    Set feeRS = New ADODB.Recordset
    feeRS.Open "charge", libCON, adOpenKeyset, adLockReadOnly
    Label12.Caption = feeRS!charge_fee
    Label13.Caption = Val(Label11.Caption) * Val(Label12.Caption)
    Label12.Caption = Format(Label12.Caption, "###.00")
    Label13.Caption = Format(Label13.Caption, "###.00")
Else
    Label11.Caption = "0"
    Label12.Caption = "0.00"
    Label13.Caption = "0.00"
End If
End Sub
