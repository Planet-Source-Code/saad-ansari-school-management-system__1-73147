VERSION 5.00
Begin VB.Form BackupDatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Utility"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6960
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6615
      Begin VB.CommandButton cmdset 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BackUp Database"
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         MouseIcon       =   "SchoolYear.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdCan 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         MouseIcon       =   "SchoolYear.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"SchoolYear.frx":02A4
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Utility"
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
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "SchoolYear.frx":0393
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7000
   End
End
Attribute VB_Name = "BackupDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCan_Click()
Unload Me
End Sub

Private Sub cmdset_Click()
On Error GoTo Err
    Shell "ntbackup.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "Error : " & " " & Error$, vbCritical + vbOKOnly, Error
    MsgBox "You don't have Backup Utility installed in your computer.", vbExclamation, "BackUp Utility Missing"
End Sub

