VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResult 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Result"
   ClientHeight    =   8655
   ClientLeft      =   1590
   ClientTop       =   1725
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   27
      Top             =   7680
      Width           =   7215
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Reset"
         Height          =   375
         Left            =   480
         MouseIcon       =   "frmResult.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         MouseIcon       =   "frmResult.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   3600
         MouseIcon       =   "frmResult.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         MouseIcon       =   "frmResult.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox stu_id 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox comRes 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmResult.frx":0548
      Left            =   2520
      List            =   "frmResult.frx":0552
      TabIndex        =   9
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox txtMax 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6240
      Width           =   2055
   End
   Begin VB.ComboBox comExamnm 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmResult.frx":0562
      Left            =   2520
      List            =   "frmResult.frx":0572
      TabIndex        =   6
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txtMarkObt 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   6480
      TabIndex        =   8
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox comSub 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   2520
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
   End
   Begin VB.ComboBox comStd 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "<<--SELECT-->>"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox comDiv 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmResult.frx":05AA
      Left            =   840
      List            =   "frmResult.frx":05AC
      TabIndex        =   2
      Text            =   "<<--SELECT-->>"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox comRoll 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   3480
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtExam 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   59768835
      CurrentDate     =   39531
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "RESULT :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1560
      TabIndex        =   25
      Top             =   6840
      Width           =   945
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " SELECT AN ENTRY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   360
      TabIndex        =   24
      Top             =   960
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Examination Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   540
      TabIndex        =   23
      Top             =   5760
      Width           =   1995
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1575
      TabIndex        =   22
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Marks Obtained :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   4680
      TabIndex        =   21
      Top             =   6240
      Width           =   1725
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Marks :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   840
      TabIndex        =   20
      Top             =   6240
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Exam Year:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   720
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Roll no :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2640
      TabIndex        =   18
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Std :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   360
      TabIndex        =   17
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Div :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label lblnm 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1200
      TabIndex        =   15
      Top             =   3120
      Width           =   60
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Result Entry :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmResult.frx":05AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2415
      Left            =   120
      Top             =   1080
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2415
      Left            =   240
      Top             =   1200
      Width           =   8415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3615
      Left            =   120
      Top             =   3840
      Width           =   8415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3615
      Left            =   240
      Top             =   3960
      Width           =   8415
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_save_Click()
If lblnm.Caption = "" Then
MsgBox "please select a student"
Exit Sub
End If

If comSub.Text = "" Then
MsgBox " please select a subject "
Exit Sub
End If

If comExamnm.Text = "" Or txtMax.Text = "" Or txtMarkObt.Text = "" Or comRes.Text = "" Then
MsgBox " please select required fields"
Exit Sub
End If
Call connect
If rs_find.State = adStateOpen Then rs_find.Close
rs_find.Open "select * from result where student_id= " & Val(stu_id.Text) & " and exam_name= '" & comExamnm.Text & "' and Std='" & ComStd.Text & "' and Div='" & ComDiv.Text & "' and roll_no= '" & comRoll.Text & "' and subject='" & comSub.Text & "' and exam_date= '" & Format(dtExam, "yyyy") & "'", con, adOpenDynamic, adLockPessimistic
If rs_find.RecordCount > 0 Then
MsgBox "marks entry already present"
rs_find.Close
Exit Sub
End If


On Error Resume Next
Dim a As Integer
a = comSub.ListCount
Label4.Caption = a
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from result", con, adOpenDynamic, adLockPessimistic
.AddNew
.Fields("student_id") = stu_id.Text
.Fields("Std") = ComStd.Text
.Fields("Div") = ComDiv.Text
.Fields("roll_no") = comRoll.Text
.Fields("exam_date") = Format(dtExam.Value, "yyyy")
.Fields("subject") = comSub.Text
.Fields("exam_name") = comExamnm.Text
.Fields("max_marks") = txtMax.Text
.Fields("marks_obt") = txtMarkObt.Text
.Fields("result") = comRes.Text
.Fields("createdBy") = SchoolMain.Label2.Caption
.Update
.Close
MsgBox "Record Successfully Entered", vbInformation + vbOKOnly, "RESULT"
comSub.RemoveItem comSub.ListIndex
End With

End Sub

Private Sub ComDiv_Click()
comRoll.Clear
comSub.Clear
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select roll_no from student_mstr where Std='" & ComStd.Text & "' and Div ='" & ComDiv.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comRoll.AddItem .Fields("roll_no").Value
.MoveNext
Loop
.Close
End With
End Sub


Private Sub comExamnm_Click()
If comExamnm.Text = "I UNIT TEST" Then
txtMax.Text = 40
ElseIf comExamnm.Text = "II UNIT TEST" Then
txtMax.Text = 40
ElseIf comExamnm.Text = "I SEMESTER" Then
txtMax.Text = 100
ElseIf comExamnm.Text = "II SEMESTER" Then
txtMax.Text = 100
End If
End Sub

Private Sub comRoll_Click()
On Error Resume Next
comSub.Clear
stu_id.Text = ""
lblnm.Caption = ""
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name,student_id from student_mstr where Std='" & ComStd.Text & "' and Div ='" & ComDiv.Text & "'and roll_no='" & comRoll.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
stu_id.Text = .Fields("student_id").Value
lblnm.Caption = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
.MoveNext
Loop
.Close
End With
With rs_find
If .State = adStateOpen Then .Close
.Open "select Subject1,Subject2,Subject3,Subject4,Subject5,Subject6,Subject7,Subject8,Subject9,Subject10,Subject11,Subject12 from class_mstr where Std = '" & ComStd.Text & "' and Div='" & ComDiv.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comSub.AddItem .Fields("Subject1")
comSub.AddItem .Fields("Subject2")
comSub.AddItem .Fields("Subject3")
comSub.AddItem .Fields("Subject4")
comSub.AddItem .Fields("Subject5")
comSub.AddItem .Fields("Subject6")
comSub.AddItem .Fields("Subject7")
comSub.AddItem .Fields("Subject8")
comSub.AddItem .Fields("Subject9")
comSub.AddItem .Fields("Subject10")
comSub.AddItem .Fields("Subject11")
comSub.AddItem .Fields("Subject12")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub ComStd_Click()
ComDiv.Clear
comSub.Clear
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select Div from class_mstr where Std='" & ComStd.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComDiv.AddItem .Fields("Div").Value
.MoveNext
Loop
.Close
End With
End Sub


Private Sub comSub_Click()
If comSub.Text = "" Then
MsgBox "please select Subject"
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call connect
dtExam.Value = Date
With rs_find
If .State = adStateOpen Then .Close
.Open "select distinct Std from class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComStd.AddItem .Fields("Std").Value
.MoveNext
Loop
.Close
End With
End Sub

Private Sub txtMarkObt_Change()
If Val(txtMax.Text) < Val(txtMarkObt.Text) Then
MsgBox "Invalid Marks Entered", vbInformation + vbOKOnly, "INVALID MARKS"
txtMarkObt.Text = ""
txtMarkObt.SetFocus
End If
End Sub


Private Sub txtMarkObt_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
