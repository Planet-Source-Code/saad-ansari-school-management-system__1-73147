VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResultRep 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Result Entry"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7260
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   480
      TabIndex        =   16
      Top             =   3600
      Width           =   6255
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
         ItemData        =   "frmResultRep.frx":0000
         Left            =   2460
         List            =   "frmResultRep.frx":0010
         TabIndex        =   4
         Text            =   "<<--SELECT-->>"
         Top             =   240
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtExam 
         Height          =   375
         Left            =   2460
         TabIndex        =   5
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
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
         Format          =   73859075
         CurrentDate     =   39531
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
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Year"
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
         Left            =   1260
         TabIndex        =   17
         Top             =   720
         Width           =   1095
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   6255
      Begin VB.ComboBox comStd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Text            =   "<<--SELECT-->>"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox comDiv 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmResultRep.frx":0048
         Left            =   960
         List            =   "frmResultRep.frx":004A
         TabIndex        =   2
         Text            =   "<<--SELECT-->>"
         Top             =   1560
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
         Left            =   3600
         Style           =   1  'Simple Combo
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
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
         Left            =   2760
         TabIndex        =   15
         Top             =   1080
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
         Left            =   480
         TabIndex        =   14
         Top             =   1080
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
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lblnm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Student Name"
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
         Height          =   360
         Left            =   1200
         TabIndex        =   12
         Top             =   360
         Width           =   4380
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
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   5520
      MouseIcon       =   "frmResultRep.frx":004C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdShow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Show Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmResultRep.frx":019E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox stu_id 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbldate 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
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
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmResultRep.frx":02F0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmResultRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdshow_Click()
Call connect
If lblnm.Caption = "Student Name" Then
MsgBox "please select an student"
Exit Sub
End If

If rs_result.State = adStateOpen Then rs_result.Close
With rs_result
.Open "select distinct subject, max_marks, marks_obt, result , exam_date from result where student_id=" & Val(stu_id.Text) & "  and exam_name = '" & comExamnm.Text & "' and exam_date= '" & Format(dtExam, "yyyy") & "'", con, adOpenStatic, adLockPessimistic
If .RecordCount <= 0 Then
            MsgBox "No Entry For This Examination !", vbExclamation, Me.Caption
            Exit Sub
End If
Set ResultReport.DataSource = rs_result
ResultReport.Sections("section4").Controls("class").Caption = ComStd.Text & " " & ComDiv.Text
ResultReport.Sections("section4").Controls("lblid").Caption = stu_id.Text
ResultReport.Sections("section4").Controls("lblname").Caption = lblnm.Caption
ResultReport.Sections("section4").Controls("lblexam_name").Caption = comExamnm.Text
ResultReport.Sections("section4").Controls("lblroll_no").Caption = comRoll.Text
ResultReport.Sections("section4").Controls("lblyear").Caption = Format(dtExam.Value, "yyyy")
.MoveFirst
 Dim max, tot, per As Integer
 max = 0
 tot = 0
 per = 0
 Do Until .EOF
                    max = max + Val(.Fields("max_marks"))
                    tot = tot + Val(.Fields("marks_obt"))
                    .MoveNext
 Loop
            per = tot / max * 100
            ResultReport.Sections("section3").Controls("lbltotal_max_marks").Caption = max
            ResultReport.Sections("section3").Controls("lbltotal_marks_obt").Caption = tot
            ResultReport.Sections("section3").Controls("lbltotal_percentage").Caption = per
ResultReport.Show
Unload Me
End With
Unload Me
End Sub

Private Sub comRoll_Click()
On Error Resume Next
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


End Sub

Private Sub ComStd_Click()
ComDiv.Clear

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

Private Sub Form_Load()
dtExam.Value = Date
Call connect
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
Private Sub ComDiv_Click()
comRoll.Clear
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

