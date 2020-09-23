VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStffGeneral 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Staff General Information"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin VB.CommandButton StffGenCls 
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdStdGenFull 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Full record         and        Edit "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      Picture         =   "frmStffGeneral.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   735
      Left            =   6480
      Picture         =   "frmStffGeneral.frx":0681
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox comstffid 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Texfname 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Texlame 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Texmname 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid FlexMember 
      Height          =   3855
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   12640511
      BackColorSel    =   255
      BackColorBkg    =   16777215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name :"
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
      Left            =   300
      TabIndex        =   13
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
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
      Left            =   480
      TabIndex        =   12
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Image pcbox 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Photograph :"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   3960
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff General Info :"
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
      Picture         =   "frmStffGeneral.frx":12C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmStffGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdStdGenFull_Click()
Me.Hide
Frmeditstaff.Show
Frmeditstaff.WindowState = vbMaximized
End Sub

Private Sub Cmdgo_Click()
On Error Resume Next
Texfname.Text = ""
Texmname.Text = ""
Texlame.Text = ""
pcbox.Picture = Nothing
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr where staff_id = " & Val(comstffid.Text), con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texfname.Text = .Fields("fname")
Texmname.Text = .Fields("mname")
Texlame.Text = .Fields("lname")
Dim a As String
a = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(a)
.MoveNext
Loop
.Close
End With
End Sub
Private Sub FlexMember_Click()
On Error Resume Next
comstffid.Text = ""
Texfname.Text = ""
Texmname.Text = ""
Texlame.Text = ""
pcbox.Picture = Nothing

Call connect
If rs_stfgrid.State = adStateOpen Then rs_stfgrid.Close
rs_stfgrid.Open "select * from staff_mstr", con, adOpenDynamic, adLockOptimistic
rs_stfgrid.MoveFirst
rs_stfgrid.Move FlexMember.Row - 1
comstffid.Text = FlexMember.TextMatrix(FlexMember.Row, 1)
Texfname.Text = FlexMember.TextMatrix(FlexMember.Row, 2)
Texmname.Text = FlexMember.TextMatrix(FlexMember.Row, 3)
Texlame.Text = FlexMember.TextMatrix(FlexMember.Row, 4)
a = App.Path & FlexMember.TextMatrix(FlexMember.Row, 40)
pcbox.Picture = LoadPicture(a)
End Sub
Public Function fillgrid()
Call connect
Dim sql As String
rs_stfgrid.CursorLocation = adUseClient
sql = "select * from staff_mstr"
rs_stfgrid.Open sql, con, adOpenKeyset, adLockOptimistic
   With FlexMember
      FlexMember.Cols = rs_stfgrid.Fields.count + 1
      FlexMember.ColWidth(0) = 0
        For c = 0 To rs_stfgrid.Fields.count - 1
          FlexMember.TextMatrix(0, c + 1) = rs_stfgrid(c).Name
        Next
      FlexMember.Rows = rs_stfgrid.RecordCount + 1
        For r = 1 To rs_stfgrid.RecordCount
           For c = 0 To rs_stfgrid.Fields.count - 1
             FlexMember.TextMatrix(r, c + 1) = IIf(IsNull(rs_stfgrid(c).Value), "", rs_stfgrid(c).Value)
           Next c
          rs_stfgrid.MoveNext
        Next r
   End With
   FlexMember.ColWidth(1) = 850
   FlexMember.ColWidth(2) = 1250
   FlexMember.ColWidth(3) = 1250
   FlexMember.ColWidth(4) = 1250
   FlexMember.ColWidth(5) = 2000
   FlexMember.ColWidth(6) = 1000
   FlexMember.ColWidth(7) = 1000
   FlexMember.ColWidth(8) = 1000
   FlexMember.ColWidth(9) = 800
   FlexMember.ColWidth(10) = 900

End Function

Private Sub Form_Load()
fillgrid
End Sub

Private Sub StffGenCls_Click()
Unload Me
End Sub
