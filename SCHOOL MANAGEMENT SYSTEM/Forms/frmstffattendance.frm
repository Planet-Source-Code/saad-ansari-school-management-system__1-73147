VERSION 5.00
Begin VB.Form frmstffattendance 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Daily Attendance"
   ClientHeight    =   4635
   ClientLeft      =   2100
   ClientTop       =   2490
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5625
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5415
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
         Height          =   1575
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
         Begin VB.CommandButton cmdok 
            BackColor       =   &H00C0E0FF&
            Caption         =   "OK"
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frmstffattendance.frx":0000
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.ComboBox comstatus 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "frmstffattendance.frx":0152
            Left            =   120
            List            =   "frmstffattendance.frx":015F
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton cmdXt 
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
            Left            =   120
            MouseIcon       =   "frmstffattendance.frx":017B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.ComboBox comname 
         BackColor       =   &H00C0FFFF&
         Height          =   2325
         Left            =   240
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Texname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1215
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
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Daily Attendance"
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
      Picture         =   "frmstffattendance.frx":02CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmstffattendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Double

Private Sub cmdok_Click()
If Texname.Text = "" Then
MsgBox "select staff_id"
Exit Sub
End If
If comname.Text = "" Then
MsgBox "select staff_id"
Exit Sub
End If
If comstatus.Text = "" Or comname.Text = "" Then
MsgBox "select status"
Exit Sub
End If
c = Val(comname.Text)
With rs_att
If .State = adStateOpen Then .Close
.Open "select * from attendance", con, adOpenDynamic, adLockPessimistic
.AddNew
.Fields("staff_id") = c
.Fields("date") = Date
.Fields("status") = comstatus.Text
.Update
.Close
End With
comname.RemoveItem comname.ListIndex
Texname.Text = ""
End Sub



Private Sub cmdxt_Click()
Unload Me
End Sub

Private Sub comname_Click()
Texname.Text = ""
Call connect
With rs_att
c = Val(comname.Text)
If .State = adStateOpen Then .Close
.Open "select fname,mname,lname from staff_mstr where staff_id = " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texname.Text = .Fields("fname") & " " & .Fields("mname") & " " & .Fields("lname")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
lbldate.Caption = Format(Date, "long date")
On Error Resume Next
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select staff_id,fname,mname,lname from staff_mstr where staff_mstr.staff_id not in ( select staff_id from attendance where date=# " & Format(Date, "m-d-yy") & "#)", con, adOpenStatic, adLockOptimistic
If Not .EOF Or .BOF Then
rs_att.MoveFirst
End If
Do Until .EOF
comname.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With
End Sub
