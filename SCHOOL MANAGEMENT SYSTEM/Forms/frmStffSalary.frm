VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStffSalary 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Staff Salary Information"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9405
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   9135
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
         Left            =   5760
         MouseIcon       =   "frmStffSalary.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_can 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   4320
         MouseIcon       =   "frmStffSalary.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Enabled         =   0   'False
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
         Left            =   2880
         MouseIcon       =   "frmStffSalary.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add"
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
         Left            =   1440
         MouseIcon       =   "frmStffSalary.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   9135
      Begin VB.TextBox Texname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox com_id 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox texsal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6120
         TabIndex        =   18
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salary Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   9135
      Begin MSComCtl2.DTPicker DTPsalYear 
         Height          =   450
         Left            =   6720
         TabIndex        =   13
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   58785795
         CurrentDate     =   40264
      End
      Begin VB.TextBox txtsal 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   " "
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmbmon 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         ItemData        =   "frmStffSalary.frx":0548
         Left            =   6720
         List            =   "frmStffSalary.frx":0570
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtlyear 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtlsal 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   1935
      End
      Begin VB.TextBox txtlmon 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Salary Paid (Rs.)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary (Rs.)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary of Month"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary of Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Salary Paid in the Month"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Salary Paid in the Year"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Salary Info :"
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
      Picture         =   "frmStffSalary.frx":05D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmStffSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_can_Click()
txtsal.Text = ""
cmbmon.Text = ""
cmbmon.Enabled = False
txtsal.Enabled = False
DTPsalYear.Enabled = False
cmdsave.Enabled = False
cmd_can.Enabled = False
cmdadd.Enabled = True

End Sub

Private Sub cmdAdd_Click()
cmbmon.Enabled = True
txtsal.Enabled = True
DTPsalYear.Enabled = True
txtsal.Text = Texsal.Text
cmdsave.Enabled = True
cmd_can.Enabled = True
cmdadd.Enabled = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from salary where staff_id=" & Val(com_id.Text) & " and salMonth='" & cmbmon.Text & "' and salYear='" & Format(DTPsalYear, "yyyy") & "'", con, adOpenDynamic, adLockOptimistic
If .RecordCount >= 1 Then
MsgBox "The Entry for the month of " & cmbmon.Text & " " & Format(DTPsalYear, "yyyy") & " is Present"
.Close
cmbmon.Text = ""
cmbmon.SetFocus
Exit Sub
End If
.Close
End With


With rs_find
If .State = adStateOpen Then .Close
.Open "select * from salary", con, adOpenDynamic, adLockOptimistic
.AddNew
.Fields("staff_id") = Val(com_id.Text)
.Fields("salPaid") = txtsal.Text
.Fields("salMonth") = cmbmon.Text
.Fields("salYear") = Format(DTPsalYear, "yyyy")
.Update
.Close
MsgBox "Salary for the month of " & cmbmon.Text & " has been Updated !"

End With
txtlmon.Text = ""
txtlyear.Text = ""
txtlsal.Text = ""
Texname.Text = ""
Texsal.Text = ""
txtsal.Text = ""
cmbmon.Text = ""
cmbmon.Enabled = False
txtsal.Enabled = False
DTPsalYear.Enabled = False
cmd_can.Enabled = False
cmdadd.Enabled = True
cmdsave.Enabled = False
End Sub

Private Sub com_id_Click()
On Error Resume Next
Texname.Text = ""
Texsal.Text = ""
txtlsal.Text = ""
txtlmon.Text = ""
txtlyear.Text = ""
txtsal.Text = ""
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr where staff_id=" & Val(com_id.Text), con, adOpenDynamic, adLockOptimistic
Do Until .EOF
Texname.Text = .Fields("lname") & " " & .Fields("fname") & " " & .Fields("mname")
Texsal.Text = .Fields("salary")
.MoveNext
Loop
.Close
End With

With rs_find
If .State = adStateOpen Then .Close
.Open "select * from salary where staff_id=" & Val(com_id.Text), con, adOpenDynamic, adLockOptimistic
If .RecordCount <= 0 Then
.Close
MsgBox "No Entry for this Record !"
Exit Sub
End If

Do Until .EOF
txtlsal.Text = .Fields("salPaid")
txtlmon.Text = .Fields("salMonth")
txtlyear.Text = .Fields("salYear")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
com_id.AddItem .Fields("staff_id")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub texsal_Change()
Texsal.Text = Format(Val(Texsal), 0#)
End Sub

