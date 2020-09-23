VERSION 5.00
Begin VB.Form frmNewDepart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Department"
   ClientHeight    =   5220
   ClientLeft      =   1320
   ClientTop       =   3000
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add &New"
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmNewDepart.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5415
      Begin VB.TextBox Texdid 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Texdn 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Texdes 
         BackColor       =   &H00C0FFFF&
         Height          =   1695
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   720
         TabIndex        =   8
         Top             =   1800
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      MouseIcon       =   "frmNewDepart.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_ok_NewDep 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Cancel"
      DownPicture     =   "frmNewDepart.frx":02A4
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmNewDepart.frx":2A138
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Department :"
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
      Picture         =   "frmNewDepart.frx":2A28A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5800
   End
End
Attribute VB_Name = "frmNewDepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_ok_NewDep_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call connect
With rs_dep
If .State = adStateOpen Then .Close
.Open "select * from department", con, adOpenDynamic, adLockPessimistic
.AddNew
.Fields("department_id") = Texdid.Text
.Fields("department_name") = Texdn.Text
.Fields("description") = Texdes.Text
.Update
.Close
End With
Command2.Enabled = True
Command1.Enabled = False
Dim rep
rep = MsgBox("Record Succesfully Created " & vbCrLf & "Would you like to create new department ?", vbInformation + vbYesNo, "RECORD CREATED")
If rep = vbYes Then
Unload Me
Call SchoolMain.mnuNewDepart_Click
Else
 Unload Me
End If
End Sub

Private Sub Command2_Click()
Texdn.Locked = False
Texdes.Locked = False
Texdn.Text = ""
Texdes.Text = ""
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Form_Load()

Call CenterForm(frmNewDepart)
On Error Resume Next
Call connect
con.Refresh
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from department", con, adOpenDynamic, adLockPessimistic
.MoveLast
If IsNull(.Fields("department_id").Value) Then
Texdid.Text = 1
Else
no = .Fields("department_id") + 1
Texdid.Text = no
End If
.Close
End With

End Sub
