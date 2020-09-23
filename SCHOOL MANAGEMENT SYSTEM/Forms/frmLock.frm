VERSION 5.00
Begin VB.Form frmLock 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   4185
   ClientTop       =   3765
   ClientWidth     =   7905
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Force Quit"
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmLock.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      MouseIcon       =   "frmLock.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtUnlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "NEW LIFE ENGLISH SCHOOL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   5790
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0FF&
      Height          =   735
      Left            =   1800
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type the correct password to UNLOCK the system"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   4815
   End
   Begin VB.Image imgopen 
      Height          =   480
      Left            =   3120
      Picture         =   "frmLock.frx":02A4
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgclose 
      Height          =   480
      Left            =   2670
      Picture         =   "frmLock.frx":06E6
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgkey 
      Height          =   315
      Left            =   3450
      Picture         =   "frmLock.frx":0B28
      Stretch         =   -1  'True
      Top             =   2865
      Width           =   330
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub forminit()
    imgkey.Left = 3120
    Timer1_Timer
End Sub

Private Sub cmdQuit_Click()
Dim rep
rep = MsgBox("Are you sure you wanna Quit ?", vbYesNo, "Force Quit")
If rep = vbYes Then
MsgBox "Session Abruptly Terminated", vbOKOnly, "Program Exit"
            Call connect
            With rs_find
            .Open "select * from UserLog", con, adOpenDynamic, adLockPessimistic
            .AddNew
            .Fields("UserID").Value = SchoolMain.Label2.Caption
            .Fields("SessionStart").Value = frmLogin.lblval.Caption
            .Fields("SessionEnd").Value = Now
            .Fields("Description") = "Session Abruptly Terminated"
            .Update
            .Close
            End With

End
Else
Exit Sub
End If
End Sub

Private Sub Command1_Click()
  Dim count As Integer
     Call connect
       If txtUnlock.Text = "" Then
       MsgBox "Please provide valid user Password !"
       Exit Sub
       End If
        Dim a, b As String
        With rs_find
        .Open "select * from user_mstr where user_id='" & SchoolMain.Label2.Caption & "'", con, adOpenDynamic, adLockPessimistic
        Do Until .EOF
        a = .Fields("pass")
        b = .Fields("user_id")
        .MoveNext
        Loop
        .Close
        End With
        If txtUnlock.Text = a And SchoolMain.Label2.Caption = b Then
            Unload Me
        Else
            count = count + 1
            MsgBox "           I N V A L I D   P A S S W O R D            " & vbCrLf _
                & "\nPlease type the correct password to unlock the system", vbCritical + vbOKOnly, "A C C E S S   D E N I E D  "
                Call Highlight(txtUnlock)
                txtUnlock.Text = ""
                txtUnlock.SetFocus
        End If
        
End Sub

Private Sub Form_Load()
Me.WindowState = vbNormal
End Sub

Private Sub Timer1_Timer()
    If imgclose.Visible = True Then
        imgclose.Visible = False
        imgopen.Visible = True
        imgkey.Left = 3585
    ElseIf imgopen.Visible = True Then
        imgopen.Visible = False
        imgclose.Visible = True
        imgkey.Left = 3930
    End If
End Sub

