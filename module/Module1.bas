Attribute VB_Name = "Module1"
'####################################################################################
'       Project Name : Employees Daily time Record System
'        Module Name : Log Form
' Module Discription : Accept and capture date and time of employees log
'        Author Name : andri_elistiawan
'              Email : andryupuk@gmail.com
'          Copyright : 2018 @ EDP, PT.Ria Indah Mandiri
'####################################################################################

Public conn As ADODB.Connection
Public LogInRS As ADODB.Recordset
Public InfoRS As ADODB.Recordset
Public SumRS As ADODB.Recordset
Public SecRS As ADODB.Recordset
Public FORMNAME As String
Public Sub DBLOAD()

    Set conn = New ADODB.Connection
        
        With conn
            .Provider = "Microsoft.jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\database\log.mdb;Persist Security Info= False"
            .Open
        End With
        
    Form1.Timer1.Enabled = True
End Sub
Public Sub LogIn()
    Set LogInRS = New ADODB.Recordset
    With LogInRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM TIMELOG"
    End With
End Sub

Public Sub Info()
    Set InfoRS = New ADODB.Recordset
    With InfoRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM INFO"
    End With
End Sub
Public Sub Sum()
    Set SumRS = New ADODB.Recordset
    With SumRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM SUMMARY"
    End With
End Sub
Public Sub Sec()
    Set SecRS = New ADODB.Recordset
    With SecRS
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM SECURITY"
    End With
End Sub

