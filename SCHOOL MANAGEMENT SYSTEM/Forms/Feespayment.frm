VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Feespayment 
   BackColor       =   &H00FFFFFF&
   Caption         =   " Fees Payment"
   ClientHeight    =   6750
   ClientLeft      =   1230
   ClientTop       =   3270
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   46
      Top             =   5880
      Width           =   9015
      Begin VB.CommandButton cmd_rst 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Reset"
         Height          =   375
         Left            =   1320
         MouseIcon       =   "Feespayment.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_can 
         BackColor       =   &H00C0E0FF&
         Caption         =   "C&ancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         MouseIcon       =   "Feespayment.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFeesPayOk 
         BackColor       =   &H00C0E0FF&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   6120
         MouseIcon       =   "Feespayment.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_chk 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         MouseIcon       =   "Feespayment.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Submit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Feespayment.frx":0548
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   9015
      Begin MSComCtl2.DTPicker DTPfees 
         Height          =   375
         Left            =   1920
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   59113473
         CurrentDate     =   40239
      End
      Begin VB.TextBox txtfproll 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   660
         Left            =   3960
         Picture         =   "Feespayment.frx":069A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtfpdiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtfpstd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TexfeepayID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtfpfname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtfpmname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtfplname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "comp_fees"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtpaid 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "comp_fees"
         Text            =   "0"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   705
         TabIndex        =   45
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll_no :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1155
         TabIndex        =   43
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4800
         TabIndex        =   40
         Top             =   4200
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4800
         TabIndex        =   39
         Top             =   3720
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4800
         TabIndex        =   38
         Top             =   3240
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Student ID:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "First name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   930
         TabIndex        =   34
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Div :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1545
         TabIndex        =   33
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Std:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1560
         TabIndex        =   32
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   750
         TabIndex        =   31
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Last name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   945
         TabIndex        =   30
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   25
         Top             =   3720
         Width           =   540
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Due :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4950
         TabIndex        =   24
         Top             =   4200
         Width           =   510
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tution Fees (Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   23
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "General Fund (Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   22
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual Charges(Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   21
         Top             =   1320
         Width           =   1950
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Examination Fee(Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   20
         Top             =   1800
         Width           =   2010
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Fee (Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   19
         Top             =   2280
         Width           =   1845
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   18
         Top             =   3240
         Width           =   600
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Fee (Rs):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4920
         TabIndex        =   17
         Top             =   2760
         Width           =   1860
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Payment:-"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Feespayment.frx":12DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9100
   End
End
Attribute VB_Name = "Feespayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a As Integer

Private Sub cmd_can_Click()
Call Command3_Click
Command2.Enabled = False
cmd_can.Enabled = False
Cmd_chk.Enabled = True
End Sub

Private Sub Cmd_chk_Click()
Cmd_chk.Enabled = False
Command2.Enabled = True
cmd_can.Enabled = True
On Error Resume Next
If Text10.Text = "" Then
Text10.Text = Val(Text9.Text) - Val(txtpaid.Text)
Else
Text10.Text = Val(Text10.Text) - Val(txtpaid.Text)
End If
End Sub

Private Sub cmd_rst_Click()
TexfeepayID.Text = ""
txtpaid.Text = 0
Call Command3_Click
End Sub

Private Sub cmdFeesPayOk_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Command2.Enabled = False
cmd_can.Enabled = False
Cmd_chk.Enabled = True

On Error Resume Next
If TexfeepayID.Text = "" Or TexfeepayID.Text = "" Or Text9.Text = "" Or txtpaid.Text = 0 Or Text10.Text = "" Then
MsgBox "Please fill the asterisk marked field", vbInformation + vbOKOnly, "Field Empty"
Else
Call connect
With rs_feepay
    If .State = adStateOpen Then .Close
    .Open "SELECT * FROM Fees_Payment", con, adOpenDynamic, adLockPessimistic
    .AddNew
    .Fields("Student_ID").Value = Feespayment.TexfeepayID.Text
    .Fields("Paid") = Feespayment.txtpaid.Text
    .Fields("Due") = Feespayment.Text10.Text
    .Fields("fees_date") = DTPfees.Value
    .Fields("fees_time") = Format(Now, "hh:mm:ss")
    .Fields("total") = Text9.Text
    .Update
    .Close
End With
MsgBox "Payment successfully accomplished", vbInformation + vbOKOnly, "Fee Payment"
End If
Dim rep
rep = MsgBox("Would you like to perform another transaction ? ", vbQuestion + vbYesNo, "New Transaction")
If rep = vbYes Then
txtpaid.Text = 0
Call Command3_Click
Else
Unload Me
End If
End Sub

Public Sub Command3_Click()
Cmd_chk.Enabled = True
On Error Resume Next
Clear
Call connect
With rs_sudfind
Dim c As Double
c = Val(TexfeepayID.Text)
If .State = adStateOpen Then .Close
.Open "select First_name,Middle_name,Last_name,Std,Div,roll_no from student_mstr where student_id=" & c & "", con, adOpenDynamic, adLockPessimistic
If .RecordCount <= 0 Then
MsgBox "no record found Enter correct Student ID"
TexfeepayID.Text = ""
TexfeepayID.SetFocus
Exit Sub
End If
Do Until .EOF
 txtfpfname.Text = .Fields("First_name").Value
 txtfpmname.Text = .Fields("Middle_name").Value
 txtfplname.Text = .Fields("Last_name").Value
 txtfpstd.Text = .Fields("Std").Value
 txtfpdiv.Text = .Fields("Div").Value
 txtfproll.Text = .Fields("roll_no").Value
 .MoveNext
 Loop
.Close
End With
With rs_feesfind
If .State = adStateOpen Then .Close
.Open "select Tution_Fees,General_Fund,Annual_Charges,Examination_Fee,Computer_Fee,Admission_Fee,Total from fees_stru where Std = '" & txtfpstd.Text & "' and Div ='" & txtfpdiv.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Text3.Text = .Fields("Tution_Fees").Value
Text4.Text = .Fields("General_Fund").Value
Text5.Text = .Fields("Annual_Charges").Value
Text6.Text = .Fields("Examination_Fee").Value
Text7.Text = .Fields("Computer_Fee").Value
Text8.Text = .Fields("Admission_Fee").Value
Text9.Text = .Fields("Total").Value
.MoveNext
Loop
.Close
End With
'Dim d As Double
'd = Val()

With rs_feepay
    If .State = adStateOpen Then .Close
    .Open "select Due from Fees_Payment where Student_ID =" & Val(Feespayment.TexfeepayID.Text), con, adOpenDynamic, adLockPessimistic
Do Until .EOF
   Text10.Text = .Fields("Due").Value
                 .MoveNext
    Loop
    .Close
End With

a = Text10.Text
End Sub

Private Sub Form_Load()
DTPfees.Value = Date
End Sub



Private Sub TexfeepayID_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Text10_Change()
If Val(Text10.Text) < 0 Then
MsgBox "Minimum value reached", vbInformation + vbOKOnly, "Error"
txtpaid.Text = ""
Text10.Text = a
End If
End Sub
Private Sub Clear()
txtfpfname.Text = ""
txtfpmname.Text = ""
txtfplname.Text = ""
txtfpstd.Text = ""
txtfpdiv.Text = ""
txtfproll.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub




Private Sub txtpaid_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
