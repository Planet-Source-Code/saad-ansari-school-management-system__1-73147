Attribute VB_Name = "ProcedureModule"
'Procedure for text highlight
Public Sub Highlight(ByRef sText As TextBox)
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub
'Procedure to center the form
Public Sub CenterForm(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = (Screen.Height - frm.Height) \ 2
    LeftCorner = (Screen.Width - frm.Width) \ 2
    frm.Move LeftCorner, TopCorner
End Sub

Public Sub Login()
frmLogin.Show
End Sub

'procedure to check the preveilege of the user
Public Sub prev()
If SchoolMain.Label8.Caption = "Administrator" Then
SchoolMain.mnuAdminn.Enabled = True
SchoolMain.mnuStaff.Enabled = True
SchoolMain.mnuDepart.Enabled = True
SchoolMain.mnuFeeStruct.Enabled = True
SchoolMain.mnuClass.Enabled = True
SchoolMain.mnustfInfo.Enabled = True
SchoolMain.mnuRepStff.Enabled = True
Else
SchoolMain.mnuAdminn.Enabled = False
SchoolMain.mnuStaff.Enabled = False
SchoolMain.mnuDepart.Enabled = False
SchoolMain.mnuFeeStruct.Enabled = False
SchoolMain.mnuClass.Enabled = False
SchoolMain.mnustfInfo.Enabled = False
SchoolMain.mnuRepStff.Enabled = False
End If
End Sub

Public Function number(KeyAscii As Integer) As Integer
Dim stat As String
stat = "0123456789"
If InStr(stat, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
KeyAscii = 0
End If
number = KeyAscii
End Function

Public Function character(KeyAscii As Integer) As Integer
Dim var  As Boolean
var = Chr(KeyAscii) Like "[a-z A-Z]"
If var = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
character = KeyAscii
End Function


Public Function uppercharacter(KeyAscii As Integer) As Integer
Dim var  As Boolean
var = Chr(KeyAscii) Like "[A-Z]"
If var = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
uppercharacter = KeyAscii
End Function
