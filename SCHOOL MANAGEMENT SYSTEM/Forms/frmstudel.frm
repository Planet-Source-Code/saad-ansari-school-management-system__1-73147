VERSION 5.00
Begin VB.Form frmstudel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Student Record"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ComDiv 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmstudel.frx":0000
      Left            =   240
      List            =   "frmstudel.frx":0002
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox ComStd 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox comroll 
      BackColor       =   &H00C0FFFF&
      Height          =   1155
      Left            =   2880
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdcls 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MouseIcon       =   "frmstudel.frx":0004
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Delete"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmstudel.frx":0156
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Texln 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Texmn 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Texfn 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox comstuid 
      Height          =   315
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   735
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
      Left            =   240
      TabIndex        =   15
      Top             =   1560
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
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label1 
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
      Left            =   2880
      TabIndex        =   13
      Top             =   720
      Width           =   570
   End
   Begin VB.Image pcbox 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Last name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1095
      TabIndex        =   11
      Top             =   3720
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Middle name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   900
      TabIndex        =   10
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "First name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Student Entry:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmstudel.frx":02A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   240
      Top             =   2280
      Width           =   6855
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   360
      Top             =   2400
      Width           =   6855
   End
End
Attribute VB_Name = "frmstudel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Double

Private Sub cmdcls_Click()
Unload Me
End Sub

Private Sub cmddel_Click()
'On Error Resume Next

If comstuid.Text = "" Then
MsgBox "please select a student"
Exit Sub
End If

Call connect
Dim rep

Dim strsql As String
strsql = "delete * from student_mstr where student_id=" & Val(comstuid.Text) & ""
rep = MsgBox("Are you sure you want to delete this record ? " & vbCrLf & "IF YOU DELETE A RECORD THE RECORD WILL BE PERMANENTLY LOST", vbExclamation + vbYesNo, "DELETE RECORD")

If rep = vbYes Then
    con.Execute strsql
    MsgBox "RECORD SUCCESSFULLY DELETED", vbInformation + vbOKOnly, "DELETE RECORD"
    comstuid.Clear
    Texfn.Text = ""
    Texln.Text = ""
    Texmn.Text = ""
    pcbox.Picture = Nothing
    comRoll.RemoveItem comRoll.ListIndex
    comstuid.Text = ""
    Else
    Exit Sub
End If

End Sub


Private Sub ComDiv_Click()
comRoll.Clear
comstuid.Clear
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select roll_no from student_mstr where Std='" & comStd.Text & "' and Div ='" & comDiv.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comRoll.AddItem .Fields("roll_no").Value
.MoveNext
Loop
.Close
End With

End Sub

Private Sub comRoll_Click()
comstuid.Clear
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where Std='" & comStd.Text & "' and Div ='" & comDiv.Text & "'and roll_no='" & comRoll.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comstuid.Text = .Fields("student_id").Value
Texfn.Text = .Fields("First_name")
Texmn.Text = .Fields("Middle_name")
Texln.Text = .Fields("Last_name")
Dim a As String
a = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(a)

.MoveNext
Loop
.Close
End With


End Sub

Private Sub comstuid_Click()
On Error Resume Next
Texfn.Text = ""
Texln.Text = ""
Texmn.Text = ""
pcbox.Picture = LoadPicture("")
Call connect
With rs_find
c = Val(comstuid.Text)
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where student_id =" & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texfn.Text = .Fields("First_name")
Texln.Text = .Fields("Last_name")
Texmn.Text = .Fields("Middle_name")

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


Private Sub ComStd_Click()
comDiv.Clear
comRoll.Clear
comstuid.Clear
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

