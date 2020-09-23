VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5310
   ClientLeft      =   2100
   ClientTop       =   3765
   ClientWidth     =   9195
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3665.056
   ScaleMode       =   0  'User
   ScaleWidth      =   8634.58
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "School INFO"
      Height          =   375
      Left            =   7200
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      MouseIcon       =   "frmAbout.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":02A4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   8676.838
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Line Line1 
      X1              =   225.372
      X2              =   8338.779
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "School Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2385
      Left            =   360
      Picture         =   "frmAbout.frx":0366
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Designed for Comprehensive Approach for Managing Activities in School."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   690
      Left            =   5520
      TabIndex        =   1
      Top             =   1920
      Width           =   3285
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION :  1.0"
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
      Height          =   345
      Left            =   6720
      TabIndex        =   3
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":36193
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmAbout.frx":362E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5400
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5520
      Top             =   1920
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmSchoolInfo.Show
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
End Sub


