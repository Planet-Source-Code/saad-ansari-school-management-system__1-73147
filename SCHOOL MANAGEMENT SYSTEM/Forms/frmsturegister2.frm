VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsturegister2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Enrollment form 2"
   ClientHeight    =   9675
   ClientLeft      =   2205
   ClientTop       =   1215
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_Bk 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Back"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8880
      Width           =   1095
   End
   Begin VB.ComboBox ComCate 
      Height          =   315
      ItemData        =   "frmsturegister2.frx":0000
      Left            =   8640
      List            =   "frmsturegister2.frx":0013
      TabIndex        =   49
      Top             =   7200
      Width           =   1455
   End
   Begin VB.ComboBox Comprogr 
      Height          =   315
      ItemData        =   "frmsturegister2.frx":0031
      Left            =   3240
      List            =   "frmsturegister2.frx":003E
      TabIndex        =   48
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ComboBox Comsclrec 
      Height          =   315
      ItemData        =   "frmsturegister2.frx":0050
      Left            =   3240
      List            =   "frmsturegister2.frx":005D
      TabIndex        =   47
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox TexCat 
      Height          =   375
      Left            =   8640
      TabIndex        =   46
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Texocc 
      Height          =   405
      Left            =   8640
      TabIndex        =   35
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Texrelation 
      Height          =   405
      Left            =   8640
      TabIndex        =   34
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Texphohome 
      Height          =   405
      Left            =   8640
      MaxLength       =   14
      TabIndex        =   33
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Texanullinc 
      Height          =   405
      Left            =   8640
      TabIndex        =   32
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Texphooffic 
      Height          =   405
      Left            =   8640
      MaxLength       =   14
      TabIndex        =   31
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Texnodepend 
      Height          =   405
      Left            =   8640
      MaxLength       =   3
      TabIndex        =   30
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Texroll 
      Height          =   405
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   29
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Texrelig 
      Height          =   405
      Left            =   8640
      TabIndex        =   22
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox Texcast 
      Height          =   405
      Left            =   8640
      TabIndex        =   21
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox Texmottong 
      Height          =   405
      Left            =   8640
      TabIndex        =   20
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8880
      Width           =   1095
   End
   Begin VB.ComboBox ComDiv 
      Height          =   315
      ItemData        =   "frmsturegister2.frx":006F
      Left            =   3240
      List            =   "frmsturegister2.frx":0071
      TabIndex        =   18
      Top             =   7800
      Width           =   1455
   End
   Begin VB.ComboBox ComStd 
      Height          =   315
      Left            =   3240
      TabIndex        =   17
      Top             =   7320
      Width           =   1455
   End
   Begin VB.ComboBox Comact 
      Height          =   315
      ItemData        =   "frmsturegister2.frx":0073
      Left            =   3240
      List            =   "frmsturegister2.frx":0080
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Texpresch 
      Height          =   405
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Texaddpresch 
      Height          =   1005
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Texclasspre 
      Height          =   405
      Left            =   3240
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "C&reate"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8880
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtyrofpass 
      Height          =   375
      Left            =   3240
      TabIndex        =   51
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   40251
   End
   Begin MSComCtl2.DTPicker dtofadm 
      Height          =   375
      Left            =   3240
      TabIndex        =   52
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   40251
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmsturegister2.frx":0097
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6720
      TabIndex        =   53
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Specify :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7200
      TabIndex        =   45
      Top             =   7680
      Width           =   1290
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Admission :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1425
      TabIndex        =   44
      Top             =   6840
      Width           =   1635
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Relationship with Student :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   6255
      TabIndex        =   43
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7455
      TabIndex        =   42
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Annual Income :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7140
      TabIndex        =   41
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Home :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7320
      TabIndex        =   40
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Office :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7320
      TabIndex        =   39
      Top             =   5040
      Width           =   1140
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parent/Gardian Information"
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
      Left            =   6240
      TabIndex        =   38
      Top             =   1920
      Width           =   3075
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Dependents in family :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   6120
      TabIndex        =   37
      Top             =   4080
      Width           =   2340
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "/Rs"
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
      Left            =   12240
      TabIndex        =   36
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Roll no :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2400
      TabIndex        =   28
      Top             =   8280
      Width           =   660
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Family Background"
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
      Left            =   6480
      TabIndex        =   23
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7680
      TabIndex        =   27
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Caste :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7860
      TabIndex        =   26
      Top             =   6360
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mother tongue :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7110
      TabIndex        =   25
      Top             =   6840
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Category :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7590
      TabIndex        =   24
      Top             =   7200
      Width           =   840
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Div :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2715
      TabIndex        =   16
      Top             =   7800
      Width           =   345
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Std. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2640
      TabIndex        =   15
      Top             =   7320
      Width           =   420
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STD  to which Admission Granted :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   6360
      Width           =   3720
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Action :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2490
      TabIndex        =   12
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Previous School Information (IF required):"
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
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   4770
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Admission :"
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
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address of previous School :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   705
      TabIndex        =   5
      Top             =   3000
      Width           =   2400
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Year of passing :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1725
      TabIndex        =   4
      Top             =   5040
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Was Promotion  Granted :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   960
      TabIndex        =   3
      Top             =   5520
      Width           =   2100
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STD in previous School :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1125
      TabIndex        =   2
      Top             =   4560
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Was the School Recognized :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   750
      TabIndex        =   1
      Top             =   4080
      Width           =   2355
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of previous School :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmsturegister2.frx":011F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   7575
      Left            =   240
      Top             =   1800
      Width           =   11055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   7575
      Left            =   480
      Top             =   1920
      Width           =   11055
   End
End
Attribute VB_Name = "frmsturegister2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pic_name As String, pic_ext As String, pic_changed As Boolean

Private Sub cmd_Bk_Click()
frmsturegister1.Show
Unload Me
End Sub

Private Sub ComCate_Click()
If ComCate.Text = "OTHER" Then
Label28.Visible = True
TexCat.Visible = True
End If
End Sub

Private Sub ComDiv_Click()
With rs_find
If .State = adStateOpen Then .Close
.Open "select Student_Strength from class_mstr where Std = '" & ComStd.Text & "' and Div = '" & ComDiv.Text & "'", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
Dim st As Integer
st = .Fields("Student_Strength")
.MoveNext
Loop
.Close
End With

With rs_find
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where Std = '" & ComStd.Text & "' and Div = '" & ComDiv.Text & "'", con, adOpenDynamic, adLockOptimistic
If .RecordCount >= st Then
MsgBox "Student Strength for Current Std and Div has Exceeded"
.Close
ComDiv.Text = ""
Texroll.Text = ""
Exit Sub
End If
End With

Texroll.Text = ""
Dim no
With rs_roll
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where Std='" & ComStd.Text & "' and Div = '" & ComDiv.Text & "'", con, adOpenDynamic, adLockPessimistic

If .EOF = False Then
.MoveLast
    no = Val(.Fields("roll_no")) + 1
    Texroll.Text = no
Else
Texroll.Text = 1
End If
.Close
End With
End Sub

Private Sub Command1_Click()

Dim cntrl As Control
For Each cntrl In Me
    If TypeOf cntrl Is TextBox Then
       If cntrl.Text = "" Then
          MsgBox "Please provide value in the blank field "
          Exit Sub
       End If
    End If
    If TypeOf cntrl Is ComboBox Then
        If cntrl.Text = "" Then
            MsgBox "Please Select the appropriate value from the dropdown box"
            Exit Sub
        End If
    End If
Next

Call connect
With rs_student

    .AddNew
    .Fields("First_name") = frmsturegister1.Texfname.Text
    .Fields("DOB") = frmsturegister1.dtdob.Value
    .Fields("Middle_name") = frmsturegister1.Texmname.Text
    .Fields("Last_name") = frmsturegister1.Texlname.Text
    .Fields("mother_name") = frmsturegister1.Texmotname.Text
    .Fields("father_name") = frmsturegister1.Texfatname.Text
    .Fields("name_of_gardian") = frmsturegister1.Texgarname.Text
    .Fields("Permanent_address") = frmsturegister1.Texaddress.Text
    .Fields("city") = frmsturegister1.Texcity.Text
    .Fields("state") = frmsturegister1.Texstate.Text
    .Fields("pincode") = frmsturegister1.Texpin.Text
    .Fields("phone_no") = frmsturegister1.Texphone.Text
    .Fields("mobile_no") = frmsturegister1.Texmobile.Text
    .Fields("Email") = frmsturegister1.Texemail.Text
    .Fields("gender") = frmsturegister1.ComGen.Text
    .Fields("place_of_birth") = frmsturegister1.TexPOB.Text
    .Fields("Nationality") = frmsturegister1.Texnaton.Text
    .Fields("Age") = frmsturegister1.Texage.Text
    .Fields("student_id") = frmsturegister1.Texstu_id.Text
    .Fields("Name_of_previous_School") = Texpresch.Text
    .Fields("Address_of_previous_School") = Texaddpresch.Text
    .Fields("STD_in_previous_School") = Texclasspre.Text
    .Fields("Year_of_passing") = dtyrofpass.Value
    .Fields("Relationship_With_student") = Texrelation.Text
    .Fields("Occupation") = Texocc.Text
    .Fields("annual_Income") = Texanullinc.Text
    .Fields("Number_of_Dependent_in_family") = Texnodepend.Text
    .Fields("Phone_Home") = Texphohome.Text
    .Fields("Phone_Office") = Texphooffic.Text
    .Fields("Religion") = Texrelig.Text
    .Fields("Caste") = Texcast.Text
    .Fields("Mother_tongue") = Texmottong.Text
    .Fields("School_mention_above_was_Recognise") = Comsclrec.Text
    .Fields("roll_no") = Texroll.Text
    .Fields("Category") = ComCate.Text
    .Fields("Std") = ComStd.Text
    .Fields("Div") = ComDiv.Text
    .Fields("Action") = Comact.Text
    .Fields("Date_of_Admission") = dtofadm.Value
    .Fields("specified_category") = TexCat.Text
    If pic_name <> "" Then
                    FileCopy pic_name, App.Path & "\Miscellaneous\STUDENT_IMAGE\" & .Fields("student_id") & pic_ext
                    .Fields("picture") = "\Miscellaneous\STUDENT_IMAGE\" & .Fields("student_id") & pic_ext

End If

    .Update
    End With
    MsgBox "Student Successfully Added, Please Proceed for Enrollment", vbInformation, "STUDENT ENROLLED"
   Dim rep
   rep = MsgBox("Do you wish to enroll new student ?", vbInformation + vbYesNo, "CREATE NEW")
   If rep = vbYes Then
       Unload frmsturegister1
       Unload Me
       frmsturegister1.Show
       
   Else
       Unload frmsturegister1
       Unload Me
   End If
End Sub

Private Sub Command2_Click()
Unload frmsturegister1
Unload Me
End Sub


Private Sub ComStd_Click()
Texroll.Text = ""
ComDiv.Clear
Call connect
With rs_class
If .State = adStateOpen Then .Close
.Open "select distinct Div from class_mstr where Std = '" & ComStd.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComDiv.AddItem .Fields("Div")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
If frmsturegister1.cdb.FileName <> "" Then
       frmsturegister1.pcbox.Picture = LoadPicture(frmsturegister1.cdb.FileName)
        pic_name = frmsturegister1.cdb.FileName
        pic_ext = Right(frmsturegister1.cdb.FileTitle, 4)
        pic_changed = True
    End If
Call connect
With rs_class
If .State = adStateOpen Then .Close
.Open "SELECT distinct Std FROM class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
ComStd.AddItem .Fields("Std")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Texanullinc_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texclasspre_Change()
If Not IsNumeric(Texclasspre) Then
MsgBox "Please provide numeric value", vbInformation, "WRONG FIELD VALUE"
Texclasspre.Text = ""
Texclasspre.SetFocus
End If
End Sub

Private Sub Texnodepend_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub


Private Sub Texphohome_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub


Private Sub Texphooffic_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texroll_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
