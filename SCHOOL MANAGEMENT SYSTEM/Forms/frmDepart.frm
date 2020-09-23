VERSION 5.00
Begin VB.Form frmDepartment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Department"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
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
      Left            =   5040
      MouseIcon       =   "frmDepart.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   3615
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   1935
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
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
         Height          =   270
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Members :"
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
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
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
         Height          =   270
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmDepart.frx":0152
         Left            =   2400
         List            =   "frmDepart.frx":0154
         MouseIcon       =   "frmDepart.frx":0156
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dept ID :"
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
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Info :"
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
      Picture         =   "frmDepart.frx":0A20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6950
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClassfindOk_Click()
Frame2.Visible = False
Unload Me
End Sub

Private Sub Combo1_Click()
Frame2.Visible = True
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Call connect
With rs_find
Dim c As Double
c = Val(Combo1.Text)
If .State = adStateOpen Then .Close
.Open "select * from department where department_id =" & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Text1.Text = .Fields("department_name")
Text3.Text = .Fields("description")
.MoveNext
Loop
.Close
End With

Dim d As Double

With rs_find
If .State = adStateOpen Then .Close
.Open "select staff_id from staff_mstr where department_id =" & Val(Combo1.Text) & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
d = d + 1
.MoveNext
Loop
.Close
End With
Text2.Text = d
End Sub
Private Sub Form_Load()
Call CenterForm(frmDepartment)
Call connect
Combo1.Clear
With rs_find
Dim c As Double
c = Val(Combo1.Text)
If .State = adStateOpen Then .Close
.Open "select * from department", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Combo1.AddItem .Fields("department_id").Value
.MoveNext
Loop
.Close
End With
End Sub



