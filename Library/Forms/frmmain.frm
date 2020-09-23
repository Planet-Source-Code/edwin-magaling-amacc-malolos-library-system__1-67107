VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "AMA Computer College - Library System"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8175
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":0442
   ScaleHeight     =   6270
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":28ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":28F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2923A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29554
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2A2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2A712
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2AB64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File Maintenance"
      Begin VB.Menu mnufilebook 
         Caption         =   "{img:1}Book Catalog"
      End
      Begin VB.Menu mnufileborrower 
         Caption         =   "{img:2}Borrower"
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilefind 
         Caption         =   "{img:3}Find"
      End
      Begin VB.Menu mnufilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "{img:4}E&xit"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "Library &Transaction"
      Begin VB.Menu mnutransborrow 
         Caption         =   "{img:5}Borrowing"
      End
      Begin VB.Menu mnutransreturn 
         Caption         =   "{img:6}Returns"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "Library &Reports"
      Begin VB.Menu mnureportsbook 
         Caption         =   "{img:7}Book's Master List"
      End
      Begin VB.Menu mnureportsborrow 
         Caption         =   "{img:7}Borrower's Master List"
      End
      Begin VB.Menu mnureportssep 
         Caption         =   "-"
      End
      Begin VB.Menu mnureportscurrent 
         Caption         =   "{img:7}Currently Borrowed Books"
      End
      Begin VB.Menu mnureportsdue 
         Caption         =   "{img:7}Due Books"
      End
      Begin VB.Menu mnureportsover 
         Caption         =   "{img:7}Overdue Books"
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "&Settings"
      Begin VB.Menu mnusettingsystem 
         Caption         =   "{img:8}System Management"
         Begin VB.Menu mnusettingsystemcategory 
            Caption         =   "Updating Book Category"
         End
         Begin VB.Menu mnusettingsystemcourse 
            Caption         =   "Updating Student Courses"
         End
         Begin VB.Menu mnusettingsystemfee 
            Caption         =   "Changing Overdue Fee"
         End
      End
      Begin VB.Menu mnusettinguser 
         Caption         =   "{img:9}User Management"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this code is exit button to prompt a message box b4 exiting
Private Const clMsgbxEXITAPP As Long = vbDefaultButton1 + vbQuestion + vbYesNo
Private mbIsDirty As Boolean

Private Sub Form_Load()
    dbconnect
    SetMenus hwnd, ImageList1
    
    'this code is exit button to prompt a message box b4 exiting
    Debug.Print "Form1::Load"
    mbIsDirty = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'this code is exit button to prompt a message box b4 exiting
    Debug.Print "Form1::QueryUnload"
    If mbIsDirty Then
        Cancel = CInt(pExitApp = False)
        If Not Cancel Then
            '-- We are ending the app. Clean up here.
            Debug.Print "Clean Up time..."
            Dim F As VB.Form
            For Each F In Forms
                Unload F
            Next
            ReleaseMenus hwnd
        End If
    End If
End Sub

Private Function pExitApp() As Boolean
'this code is exit button to prompt a message box b4 exiting
    Debug.Print "Exit Application"
    pExitApp = (MsgBox("Exit system?", clMsgbxEXITAPP, "Library") = vbYes)
End Function

Private Sub Form_Unload(Cancel As Integer)
'this code is exit button to prompt a message box b4 exiting
    Debug.Print "Form1::Unload"
End Sub

Private Sub mnuabout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnufilebook_Click()
    frmbookcatalog.Show vbModal
End Sub

Private Sub mnufileborrowerstud_Click()

End Sub

Private Sub mnufileborrower_Click()
    frmborrower.Show vbModal
End Sub

Private Sub mnufileexit_Click()
If MsgBox("Exit system?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
ReleaseMenus hwnd
End
End Sub

Private Sub mnufilefind_Click()
    frmfind.Show vbModal
End Sub

Private Sub mnureportsbook_Click()
    Set bookRS = New ADODB.Recordset
    bookRS.Open "book_catalog", libCON, adOpenKeyset, adLockReadOnly
    Set rptbook.DataSource = bookRS
    rptbook.Show vbModal
End Sub

Private Sub mnureportsborrow_Click()
    Set borrowerRS = New ADODB.Recordset
    borrowerRS.Open "borrower_record", libCON, adOpenKeyset, adLockReadOnly
    Set rptborrower.DataSource = borrowerRS
    rptborrower.Show vbModal
End Sub

Private Sub mnureportscurrent_Click()
    Set currentRS = New ADODB.Recordset
    currentRS.Open "current_borrow", libCON, adOpenKeyset, adLockReadOnly
    Set rptcurrent.DataSource = currentRS
    rptcurrent.Show
End Sub

Private Sub mnureportsdue_Click()
    Set currentRS = New ADODB.Recordset
    SQLstr = "select * from current_borrow where due_date = date()"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    Set rptdue.DataSource = currentRS
    rptdue.Show vbModal
End Sub

Private Sub mnureportsover_Click()
    Set currentRS = New ADODB.Recordset
    SQLstr = "select * from current_borrow where due_date < date()"
    currentRS.Open SQLstr, libCON, adOpenKeyset, adLockReadOnly
    Set rptoverdue.DataSource = currentRS
    rptoverdue.Show vbModal
End Sub

Private Sub mnusettingsystemcategory_Click()
    frmcategory.Show vbModal
End Sub

Private Sub mnusettingsystemcourse_Click()
    frmcourse.Show vbModal
End Sub

Private Sub mnusettingsystemfee_Click()
    frmsystem.Show vbModal
End Sub

Private Sub mnusettinguser_Click()
    frmusers.Show vbModal
End Sub

Private Sub mnutransborrow_Click()
    frmborrowing.Show vbModal
End Sub

Private Sub mnutransreturn_Click()
    frmreturning.Show vbModal
End Sub
