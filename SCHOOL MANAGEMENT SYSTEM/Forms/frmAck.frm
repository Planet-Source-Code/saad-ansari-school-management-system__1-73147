VERSION 5.00
Begin VB.Form frmAck 
   BackColor       =   &H00000000&
   Caption         =   "Acknowledgement"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Frame1"
      Height          =   11640
      Left            =   1320
      TabIndex        =   2
      Top             =   10680
      Width           =   6855
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   $"frmAck.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   11415
         Left            =   240
         TabIndex        =   3
         Top             =   -120
         Width           =   6495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11160
      Top             =   7440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   10440
      MouseIcon       =   "frmAck.frx":0C43
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A C K N O W L E D G E M E N T"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   8535
      Left            =   11400
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmAck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
Timer1.Enabled = True
a = 1
End Sub

Private Sub Timer1_Timer()
a = a + 1
Frame1.Top = Frame1.Top - 10
If Frame1.Top = -10080 Then
Frame1.Top = 9240
End If
End Sub
