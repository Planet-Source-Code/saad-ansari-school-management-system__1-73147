VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3135
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1852.262
   ScaleMode       =   0  'User
   ScaleWidth      =   6112.536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtuid 
      BackColor       =   &H00C0FFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4080
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      ToolTipText     =   "User ID"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H00C0FFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4080
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Password"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Login 
      BackColor       =   &H00C0E0FF&
      Caption         =   "LOG - IN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmLogin.frx":35E5
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Log In"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      MousePointer    =   1  'Arrow
      ScaleHeight     =   315
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox adoadmin 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      MousePointer    =   1  'Arrow
      ScaleHeight     =   315
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "passwordADmin"
      DataSource      =   "adoadmin"
      Height          =   285
      Left            =   7320
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmd_CanLog 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      MouseIcon       =   "frmLogin.frx":3737
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For any assistance Please contact your System Administrator"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   3120
      TabIndex        =   12
      Top             =   2520
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "For use  by Authorized Personnel Only....                  "
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      MouseIcon       =   "frmLogin.frx":3889
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":3B93
      ToolTipText     =   "Closes The Application"
      Top             =   -120
      Width           =   480
   End
   Begin VB.Label lblval 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Login Detail :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblUserif 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3075
      TabIndex        =   6
      Top             =   600
      Width           =   990
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   1305
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Login_Click()
On Error Resume Next
Call connect
If txtuid.Text = "" Then
MsgBox "Please Enter a Valid User ID"
txtuid.SetFocus
Exit Sub
End If
If txtpassword.Text = "" Then
MsgBox "Password field cannot be blank"
txtpassword.SetFocus
Exit Sub
Exit Sub
End If
Do Until rs_userid.EOF
If (txtuid.Text = rs_userid.Fields("user_id") And txtpassword.Text = rs_userid.Fields("pass")) Or (txtuid.Text = "Administrator" And txtpassword.Text = "admin") Then
Me.Hide
SchoolMain.Show
SchoolMain.WindowState = vbMaximized
SchoolMain.mnuAdminn.Enabled = True
With SchoolMain
    .Enabled = True
    .mnuAdminn = True
    .mnulogout.Enabled = True
    .mnuquit.Enabled = True
    .mnuView.Enabled = True
    .mnuUtil.Enabled = True
    .mnuAdminn.Enabled = True
    .mnuLock.Enabled = True
    .mnuTrans.Enabled = True
    .mnuReport.Enabled = True
    .lvMenu.Enabled = True
    .cmdChgPass.Enabled = True
    .cmdLogf.Enabled = True
End With
SchoolMain.mnulogout.Caption = "Logout"
MsgBox "WELCOME TO NLES AUTOMATED SYSTEM" & vbCrLf & vbCrLf & "  SESSION STARTS AT : " & Time, vbInformation + vbOKOnly, "SUCCESSFULLY ACCESSED"
SchoolMain.StatusBar1.Panels(1).Text = SchoolMain.StatusBar1.Panels(1).Text + " Logged In as " & txtuid.Text
'shows the field values of "Logged on" frame and "Today" frame on the MDIform

If txtuid = "Administrator" Then
SchoolMain.cmdChgPass.Enabled = False
SchoolMain.Label8.Caption = "Administrator"
End If

SchoolMain.Label2.Caption = txtuid.Text
SchoolMain.Label5.Caption = Format(Now, "ddd,mmmm d,yyyy")
SchoolMain.Label6.Caption = Format(Now, "hh:mm:ss")
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from user_mstr where user_id='" & txtuid.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
SchoolMain.Label8.Caption = .Fields("Acct_typ")
.MoveNext
Loop
.Close
End With
SchoolMain.Timer1.Enabled = True
lblval.Caption = Now
Call prev
Me.Hide
Exit Sub
End If
rs_userid.MoveNext
Loop
MsgBox "Please provide valid Id / password.", vbCritical + vbOKOnly, "ACCESS DENIED"
txtuid.SetFocus
Call Highlight(txtuid)
txtpassword.Text = ""
End Sub

Private Sub cmd_CanLog_Click()
Me.Hide
With SchoolMain
    .Enabled = True
    .mnuAdminn.Enabled = False
    .mnuLock.Enabled = False
    .mnuTrans.Enabled = False
    .mnuView.Enabled = False
    .mnuReport.Enabled = False
    .mnuUtil.Enabled = False
    .cmdLogf.Enabled = False
    .cmdChgPass.Enabled = False
    .lvMenu.Enabled = False
End With
SchoolMain.mnulogout.Caption = "Login"
End Sub

Private Sub Form_Load()
Load SchoolMain
SchoolMain.Enabled = False
End Sub

Private Sub Image1_Click()
End
End Sub



Private Sub txtpassword_GotFocus()
Call Highlight(txtpassword)
End Sub

Private Sub txtuid_GotFocus()
Call Highlight(txtuid)
End Sub
