VERSION 5.00
Begin VB.Form frmIndfees 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Fees Info"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   6960
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   6015
      Begin VB.TextBox lblnm 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   2775
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
         MouseIcon       =   "frmIndfees.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox comDiv 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmIndfees.frx":08CA
         Left            =   840
         List            =   "frmIndfees.frx":08CC
         MouseIcon       =   "frmIndfees.frx":08CE
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox comStd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   840
         MouseIcon       =   "frmIndfees.frx":1198
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   960
         Width           =   1455
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   645
         WordWrap        =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   465
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   465
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   810
      End
   End
   Begin VB.TextBox stu_id 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      MouseIcon       =   "frmIndfees.frx":1A62
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Cmdgo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      MouseIcon       =   "frmIndfees.frx":1BB4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Fees Record :"
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
      Picture         =   "frmIndfees.frx":1D06
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmIndfees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Cmdgo_Click()
Call connect

With rs_result
If .State = adStateOpen Then .Close
.Open "select * from Fees_Payment where student_id=" & Val(stu_id.Text), con, adOpenStatic, adLockPessimistic
If .RecordCount <= 0 Then
.Close
MsgBox " no fees record for current student"
Exit Sub
End If
.Close
End With


If rs_result.State = adStateOpen Then rs_result.Close
With rs_result
If .State = adStateOpen Then .Close
.Open "select * from Fees_Payment where student_id=" & Val(stu_id.Text), con, adOpenStatic, adLockPessimistic
Set IndFees.DataSource = rs_result
IndFees.Sections("section4").Controls("lblclass").Caption = ComStd.Text & " " & ComDiv.Text
IndFees.Sections("section4").Controls("lblid").Caption = stu_id.Text
IndFees.Sections("section4").Controls("lblname").Caption = lblnm.Text
IndFees.Sections("section4").Controls("lblroll_no").Caption = comRoll.Text
IndFees.Show
End With
Unload Me
End Sub

Private Sub comRoll_Click()
On Error Resume Next
stu_id.Text = ""
lblnm.Text = ""
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name,student_id from student_mstr where Std='" & ComStd.Text & "' and Div ='" & ComDiv.Text & "'and roll_no='" & comRoll.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
stu_id.Text = .Fields("student_id").Value
lblnm.Text = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
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
