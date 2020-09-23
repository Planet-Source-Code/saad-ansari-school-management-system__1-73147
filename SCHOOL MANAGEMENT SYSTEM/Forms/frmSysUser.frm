VERSION 5.00
Begin VB.Form frmSysUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Info"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmddel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Delete User"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   5520
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Encrypt Password"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Decrypt Password"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add New User"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "USER DETAILS "
      Height          =   2415
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   4935
      Begin VB.TextBox txtacc 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtstffid 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtun 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.ComboBox comuser 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select User :"
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
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "System User Administration"
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
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmSysUser.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDel_Click()
Call connect
Dim str As String

str = "delete * from user_mstr where user_id='" & comuser.Text & "'"
con.Execute str
MsgBox "User Successfully Deleted"
comuser.RemoveItem comuser.ListIndex
txtun.Text = ""
txtPass.Text = ""
txtstffid.Text = ""
txtacc.Text = ""

End Sub

Private Sub comuser_Click()
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from user_mstr where user_id='" & comuser.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
txtun.Text = .Fields("user_id")
txtPass.Text = .Fields("pass")
txtstffid.Text = .Fields("staff_id")
txtacc.Text = .Fields("Acct_typ")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
frmUserAdd.Show
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
If SchoolMain.Label8.Caption = "Administrator" Then
Frame2.Enabled = True
Command2.Enabled = True
Else
Frame2.Enabled = False
Command2.Enabled = False
End If
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from user_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comuser.AddItem .Fields("user_id")
.MoveNext
Loop
.Close
End With

End Sub


Private Sub Option1_Click(Index As Integer)
If Option1(1).Value = True Then
txtPass.PasswordChar = ""
Else
txtPass.PasswordChar = "*"
End If
End Sub

