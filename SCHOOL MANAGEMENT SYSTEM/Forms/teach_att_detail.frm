VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form teach_att_detail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Attendance"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   7455
      Begin VB.ComboBox Comstaff_a 
         BackColor       =   &H00C0FFFF&
         Height          =   2130
         Left            =   2640
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox comstaff_p 
         BackColor       =   &H00C0FFFF&
         Height          =   2130
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Comstaff_l 
         BackColor       =   &H00C0FFFF&
         Height          =   2130
         Left            =   5160
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID  present"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID absent"
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
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID on leave"
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
         Left            =   5640
         TabIndex        =   13
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7455
      Begin VB.CommandButton Cmdfind 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   375
         Left            =   5160
         MouseIcon       =   "teach_att_detail.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPdate 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   107544576
         CurrentDate     =   40245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date :"
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
         Left            =   630
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Lblname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   4095
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
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   555
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      MouseIcon       =   "teach_att_detail.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note : After Selecting Date Value Please Press Go Button               or press Enter Key."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Attendance Detail"
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
      Picture         =   "teach_att_detail.frx":02A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "teach_att_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Double

Private Sub Cmdfind_Click()
comstaff_p.Clear
Comstaff_a.Clear
Comstaff_l.Clear
Lblname.Caption = ""
Call connect
With rs_att
If .State = adStateOpen Then .Close
.Open "select staff_id from attendance where date= #" & DTPdate.Value & "# and status='PRESENT'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comstaff_p.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With
With rs_att
If .State = adStateOpen Then .Close
.Open "select staff_id from attendance where date= #" & DTPdate.Value & "# and status='ABSENT'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Comstaff_a.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With
With rs_att
If .State = adStateOpen Then .Close
.Open "select staff_id from attendance where date= #" & DTPdate.Value & "# and status='LEAVE'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Comstaff_l.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Comstaff_a_Click()
Lblname.Caption = ""
Call connect
With rs_att
c = Val(Comstaff_a.Text)
If .State = adStateOpen Then .Close
.Open "select fname,mname,lname from staff_mstr where staff_id= " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Lblname.Caption = .Fields("fname") & " " & .Fields("mname") & " " & .Fields("lname")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub Comstaff_l_Click()
Lblname.Caption = ""
Call connect
With rs_att
c = Val(Comstaff_l.Text)
If .State = adStateOpen Then .Close
.Open "select fname,mname,lname from staff_mstr where staff_id= " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Lblname.Caption = .Fields("fname") & " " & .Fields("mname") & " " & .Fields("lname")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub comstaff_p_Click()
Lblname.Caption = ""
Call connect
With rs_att
c = Val(comstaff_p.Text)
If .State = adStateOpen Then .Close
.Open "select fname,mname,lname from staff_mstr where staff_id= " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Lblname.Caption = .Fields("fname") & " " & .Fields("mname") & " " & .Fields("lname")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub DTPdate_Change()
comstaff_p.Clear
Comstaff_a.Clear
Comstaff_l.Clear
Lblname.Caption = ""
End Sub
