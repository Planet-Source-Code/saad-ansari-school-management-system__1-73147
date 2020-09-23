VERSION 5.00
Begin VB.Form frm_stu_attendance 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   1575
   ClientTop       =   2745
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox ComDiv 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmStffProfessional.frx":0000
         Left            =   600
         List            =   "frmStffProfessional.frx":0002
         TabIndex        =   10
         Text            =   "<<--SELECT-->>"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox ComStd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Text            =   "<<--SELECT-->>"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Texname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox comroll 
         BackColor       =   &H00C0FFFF&
         Height          =   2325
         Left            =   2160
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "<--SELECT ROLL_NO-->"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Frame frmall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attendance Status :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   4560
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
         Begin VB.ComboBox comstatus 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "frmStffProfessional.frx":0004
            Left            =   240
            List            =   "frmStffProfessional.frx":0011
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdok 
            BackColor       =   &H00C0E0FF&
            Caption         =   "OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmStffProfessional.frx":002D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdClassfindOk 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frmStffProfessional.frx":017F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.TextBox Tex_id 
         Height          =   375
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   855
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
         Left            =   195
         TabIndex        =   14
         Top             =   1680
         Width           =   345
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
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll no"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   570
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Daily Attendance :"
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
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -120
      Picture         =   "frmStffProfessional.frx":02D1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frm_stu_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClassfindOk_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Call connect
If comRoll.Text = "" Or comRoll.Text = "<--SELECT ROLL_NO-->" Then
MsgBox "select roll_no"
Exit Sub
End If
If comstatus.Text = "" Then
MsgBox "select status"
Exit Sub
End If
With rs_att
If .State = adStateOpen Then .Close
.Open "select * from stu_att", con, adOpenDynamic, adLockPessimistic
.AddNew
.Fields("student_id") = Tex_id.Text
.Fields("date") = Date
.Fields("status") = comstatus.Text
.Update
.Close
End With
comRoll.RemoveItem comRoll.ListIndex
Texname.Text = ""
Tex_id.Text = ""
End Sub



Private Sub ComDiv_Click()
comRoll.Clear
Tex_id.Text = ""
Texname.Text = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select student_id,roll_no from student_mstr where Std='" & comStd.Text & "' and Div ='" & comDiv.Text & "' and student_id not in (select student_id from stu_att where date=#" & Date & "#)", con, adOpenStatic, adLockOptimistic
Do Until .EOF
comRoll.AddItem .Fields("roll_no")
.MoveNext
Loop
.Close
End With

End Sub



Private Sub comRoll_Click()
Call connect
Tex_id.Text = ""
Texname.Text = ""
With rs_att
If .State = adStateOpen Then .Close
.Open "select student_id,First_name,Middle_name,Last_name from student_mstr where Std = '" & comStd.Text & "' and Div= '" & comDiv.Text & "' and roll_no= '" & comRoll.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Tex_id.Text = .Fields("student_id")
Texname.Text = .Fields("First_name") & " " & .Fields("Middle_name") & " " & .Fields("Last_name")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub ComStd_Click()
comDiv.Clear
Tex_id.Text = ""
Texname.Text = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select distinct Div from class_mstr where Std = '" & comStd.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comDiv.AddItem .Fields("Div")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "SELECT distinct Std FROM class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comStd.AddItem .Fields("Std")
.MoveNext
Loop
.Close
End With
End Sub
