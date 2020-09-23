Attribute VB_Name = "modGlobal"
Public libCON As ADODB.Connection
Public borrowerRS As ADODB.Recordset
Public catRS As ADODB.Recordset
Public bookRS As ADODB.Recordset
Public borrowRS As ADODB.Recordset
Public currentRS As ADODB.Recordset
Public courseRS As ADODB.Recordset
Public returnRS As ADODB.Recordset
Public feeRS As ADODB.Recordset
Public userRS As ADODB.Recordset
Public returnCMD As ADODB.Command
Public bookCMD As ADODB.Command
Public borrowerCMD As ADODB.Command
Public currentCMD As ADODB.Command
Public feeCMD As ADODB.Command
Public userCMD As ADODB.Command
Public courseCMD As ADODB.Command
Public catCMD As ADODB.Command
Public SQLstr As String
Public auto As Boolean

Sub main()
    frmLogin.Show
End Sub

Sub dbconnect()
    Set libCON = New ADODB.Connection
    libCON.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Databases\lib_db.mdb;Persist Security Info=False"
End Sub

