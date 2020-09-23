VERSION 5.00
Begin VB.Form teachsuspend 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff Suspension"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox comstffid 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   2430
   End
   Begin VB.TextBox Texfn 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Texln 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Texmn 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Suspend"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdcal 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image pcbox 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Middle Name:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name:"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Suspention :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "teachsuspend.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   120
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   240
      Top             =   1560
      Width           =   6975
   End
End
Attribute VB_Name = "teachsuspend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Double

Private Sub cmdcal_Click()
Unload Me
End Sub

Private Sub cmddel_Click()
Call connect
If comstffid.Text = "" Then
MsgBox " select staff id"
Exit Sub
End If


If comstffid.Text = "0" Then
MsgBox " cannot delete system administrator"
comstffid.Text = ""
Exit Sub
End If

Dim rep
Dim strsql As String
strsql = "delete * from staff_mstr where staff_id=" & c & ""
rep = MsgBox("Are you sure you want to delete this record ? " & vbCrLf & "IF YOU DELETE A RECORD THE RECORD WILL BE PERMANENTLY LOST", vbExclamation + vbYesNo, "DELETE RECORD")
If rep = vbYes Then
con.Execute strsql
MsgBox "RECORD SUCCESSFULLY DELETED", vbInformation + vbOKOnly, "DELETE RECORD"
comstffid.Clear
Texfn.Text = ""
Texln.Text = ""
Texmn.Text = ""
pcbox.Picture = Nothing
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr order by staff_id desc", con, adOpenDynamic, adLockPessimistic
.MoveFirst
Do Until .EOF
comstffid.AddItem .Fields("staff_id").Value
.MoveNext
Loop
.Close
End With
Else
Exit Sub
End If
End Sub

Private Sub comstffid_Click()
On Error Resume Next
Texfn.Text = ""
Texln.Text = ""
Texmn.Text = ""
pcbox.Picture = LoadPicture("")
Call connect
With rs_find
c = Val(comstffid.Text)
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr where staff_id =" & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texfn.Text = .Fields("fname")
Texln.Text = .Fields("lname")
Texmn.Text = .Fields("mname")
Dim a As String
a = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(a)
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr order by staff_id desc", con, adOpenDynamic, adLockPessimistic
.MoveFirst
Do Until .EOF
comstffid.AddItem .Fields("staff_id").Value
.MoveNext
Loop
.Close
End With
End Sub

