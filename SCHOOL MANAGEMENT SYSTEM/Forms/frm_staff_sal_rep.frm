VERSION 5.00
Begin VB.Form frm_staff_sal_rep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Salary"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      Begin VB.ComboBox cbostffId 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID :"
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
         Height          =   345
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff salary Info :"
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
      Left            =   0
      Picture         =   "frm_staff_sal_rep.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frm_staff_sal_rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbostffId_Click()
Unload staff_sal
txtname.Text = ""
On Error Resume Next
Call connect

With rs_find
If .State = adStateOpen Then .Close
.Open "select fname,mname,lname from staff_mstr where staff_mstr.staff_id = " & Val(cbostffId.Text) & "", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
txtname.Text = .Fields("fname") & "  " & .Fields("lname") & " " & .Fields("mname")
.MoveNext
Loop
.Close
End With


With rs_find
If .State = adStateOpen Then .Close
.Open "select salPaid,salMonth,salYear from salary where staff_id = " & Val(cbostffId.Text) & "", con, adOpenStatic, adLockOptimistic



If .RecordCount <= 0 Then
            MsgBox "No Entry For This Examination !", vbExclamation, Me.Caption
            Exit Sub
End If
Set staff_sal.DataSource = rs_find
staff_sal.Sections("section4").Controls("staffid").Caption = cbostffId.Text
staff_sal.Sections("section4").Controls("name").Caption = txtname.Text
staff_sal.Show
Unload Me
End With

Unload Me

End Sub

Private Sub Form_Load()
txtname.Text = ""

Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
cbostffId.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With
End Sub

