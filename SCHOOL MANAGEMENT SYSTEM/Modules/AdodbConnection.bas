Attribute VB_Name = "AdodbConnection"
'DECLARATION AS USED IN THE PROJECT FOR CONNECTION

Public con As New ADODB.Connection
Public adors As New ADODB.Connection
Public rs_userid As New ADODB.Recordset
Public rs_student As New ADODB.Recordset
Public rs_class As New ADODB.Recordset
Public rs_user As New ADODB.Recordset
Public rs_feepay As New ADODB.Recordset
Public rs_feestruct As New ADODB.Recordset
Public rs_sudfind As New ADODB.Recordset
Public rs_feesfind As New ADODB.Recordset
Public rs_feesstrufind As New ADODB.Recordset
Public rs_find As New ADODB.Recordset
Public rs_dep As New ADODB.Recordset
Public rs_att As New ADODB.Recordset
Public rs_staff As New ADODB.Recordset
Public rs_result As New ADODB.Recordset
Public rs_roll As New ADODB.Recordset
Public rs_classrep As New ADODB.Recordset
Public rs_stugrid As New ADODB.Recordset
Public rs_stfgrid As New ADODB.Recordset
Dim str As String

Public Sub connect()
'SUB FOR CREATING CONNECTION

Set con = New ADODB.Connection
Set rs_userid = New ADODB.Recordset
Set rs_student = New ADODB.Recordset
Set rs_class = New ADODB.Recordset
Set rs_user = New ADODB.Recordset
Set rs_feepay = New ADODB.Recordset
Set rs_feestruct = New ADODB.Recordset
Set rs_sudfind = New ADODB.Recordset
Set rs_feesfind = New ADODB.Recordset
Set rs_feesstrufind = New ADODB.Recordset
Set rs_find = New ADODB.Recordset
Set rs_dep = New ADODB.Recordset
Set rs_att = New ADODB.Recordset
Set rs_staff = New ADODB.Recordset
Set rs_result = New ADODB.Recordset
Set rs_roll = New ADODB.Recordset
Set rs_classrep = New ADODB.Recordset
Set rs_stugrid = New ADODB.Recordset
Set rs_stfgrid = New ADODB.Recordset

con.CursorLocation = adUseClient
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\Database\sms.mdb;Persist Security Info=False"
con.Open
rs_userid.Open "SELECT * FROM user_mstr", con, adOpenStatic, adLockPessimistic
rs_student.Open "SELECT * FROM student_mstr", con, adOpenStatic, adLockPessimistic
rs_user.Open "SELECT * FROM User_Mstr", con, adOpenStatic, adLockPessimistic
rs_feestruct.Open "SELECT * FROM fees_stru", con, adOpenStatic, adLockPessimistic

End Sub


'SUB TO PROVIDE DATA SOURCE TO REPORTS OF user
Sub Init_user_Report()
On Error Resume Next
    
    Dim adousers As New ADODB.Recordset
    If adousers.State = adStateOpen Then adousers.Close
    adousers.Open "user_mstr", con, adOpenStatic
    Set druser.DataSource = adousers
    druser.Refresh
    adousers.Close
End Sub

'SUB TO PROVIDE DATA SOURCE TO REPORT OF STUDENT
Sub Init_student_Report()
On Error Resume Next
    
    Dim adousers As New ADODB.Recordset
    If adousers.State = adStateOpen Then adousers.Close
    adousers.Open "student_mstr", con, adOpenStatic
    Set StudentReport.DataSource = adousers
    StudentReport.Refresh
    adousers.Close
End Sub

'SUB TO PROVIDE DATA SOURCE TO REPORT OF STAFF
Sub Init_staff_Report()
On Error Resume Next
    
    Dim adousers As New ADODB.Recordset
    If adousers.State = adStateOpen Then adousers.Close
    adousers.Open "staff_mstr", con, adOpenStatic
    Set StaffReport.DataSource = adousers
    StaffReport.Refresh
    adousers.Close
End Sub

'SUB TO PROVIDE DATA SOURCE TO REPORT OF STUDENT FEES
Sub Init_studFees_Report()
On Error Resume Next
    
    Dim adousers As New ADODB.Recordset
    If adousers.State = adStateOpen Then adousers.Close
    adousers.Open "Fees_Payment", con, adOpenStatic
    Set StudentFees.DataSource = adousers
    StudentFees.Refresh
    adousers.Close
End Sub

'SUB TO PROVIDE DATA SOURCE TO REPORT OF STUDENT RESULT
Sub Init_RESULT_Report()
On Error Resume Next
    
    Dim adousers As New ADODB.Recordset
    If adousers.State = adStateOpen Then adousers.Close
    adousers.Open "result", con, adOpenStatic
    Set ResultReport.DataSource = adousers
    ResultReport.Refresh
    adousers.Close
End Sub

