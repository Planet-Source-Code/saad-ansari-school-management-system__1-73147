VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Stu_att_detail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Attendance"
   ClientHeight    =   6780
   ClientLeft      =   600
   ClientTop       =   -930
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   7455
      Begin VB.ComboBox comroll_p 
         BackColor       =   &H00C0FFFF&
         Height          =   2715
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   15
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Comroll_a 
         BackColor       =   &H00C0FFFF&
         Height          =   2715
         Left            =   2640
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Comroll_l 
         BackColor       =   &H00C0FFFF&
         Height          =   2715
         Left            =   5160
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll no present"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll no absent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll no leave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7455
      Begin VB.ComboBox ComStd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Text            =   "<<--SELECT-->>"
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox ComDiv 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Stu_att_detail.frx":0000
         Left            =   3360
         List            =   "Stu_att_detail.frx":0002
         TabIndex        =   6
         Text            =   "<<--SELECT-->>"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Texname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPdate 
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59441153
         CurrentDate     =   40247
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Std. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2835
         TabIndex        =   11
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Div :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2835
         TabIndex        =   10
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5160
         TabIndex        =   8
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "close"
      Height          =   375
      Left            =   6120
      MouseIcon       =   "Stu_att_detail.frx":0004
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note : After Selecting Date Value Please re-select the Std.             and Div. Fields In order for the database to refresh."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Width           =   4215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Attendance Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Stu_att_detail.frx":0156
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Stu_att_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ComDiv_Click()
Texname.Text = ""
comroll_p.Clear
Comroll_a.Clear
Comroll_l.Clear
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select roll_no from stu_att,student_mstr where stu_att.status ='PRESENT'and stu_att.date= #" & DTPdate.Value & "# and Std= '" & ComStd.Text & "' and Div = '" & ComDiv.Text & "' and stu_att.student_id = student_mstr.student_id", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
comroll_p.AddItem .Fields("roll_no")
.MoveNext
Loop
.Close
End With
With rs_att
If .State = adStateOpen Then .Close
.Open "select roll_no from stu_att,student_mstr where stu_att.status ='ABSENT'and stu_att.date= #" & DTPdate.Value & "# and Std= '" & ComStd.Text & "' and Div = '" & ComDiv.Text & "' and stu_att.student_id = student_mstr.student_id", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
Comroll_a.AddItem .Fields("roll_no")
.MoveNext
Loop
.Close
End With
With rs_att
If .State = adStateOpen Then .Close
.Open "select roll_no from stu_att,student_mstr where stu_att.status ='LEAVE'and stu_att.date= #" & DTPdate.Value & "# and Std= '" & ComStd.Text & "' and Div = '" & ComDiv.Text & "' and stu_att.student_id = student_mstr.student_id", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
Comroll_l.AddItem .Fields("roll_no")
.MoveNext
Loop
.Close
End With

End Sub


Private Sub Command1_Click()
Unload Me
End Sub





Private Sub Comroll_a_Click()
Texname.Text = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name from student_mstr where Std='" & ComStd.Text & "' and Div='" & ComDiv.Text & "' and roll_no='" & Comroll_a.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texname.Text = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub Comroll_l_Click()
Texname.Text = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name from student_mstr where Std='" & ComStd.Text & "' and Div='" & ComDiv.Text & "' and roll_no='" & Comroll_l.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texname.Text = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub comroll_p_Click()
Texname.Text = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name from student_mstr where Std='" & ComStd.Text & "' and Div='" & ComDiv.Text & "' and roll_no='" & comroll_p.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texname.Text = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub ComStd_Click()
ComDiv.Clear
Texname.Text = ""
comroll_p.Clear
Comroll_a.Clear
Comroll_l.Clear
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select distinct Div from class_mstr where Std = '" & ComStd.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComDiv.AddItem .Fields("Div")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub DTPdate_Change()
Texname.Text = ""
comroll_p.Clear
Comroll_a.Clear
Comroll_l.Clear
comroll_p.Text = ""
Comroll_a.Text = ""
Comroll_l.Text = ""
End Sub

Private Sub Form_Load()
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "SELECT distinct Std FROM class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComStd.AddItem .Fields("Std")
.MoveNext
Loop
.Close
End With

End Sub

