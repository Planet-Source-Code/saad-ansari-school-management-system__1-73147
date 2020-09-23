VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudGeneral 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student General Information"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   8895
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   8655
      Begin VB.ComboBox ComStd 
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Text            =   "<<--SELECT-->>"
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox ComDiv 
         Height          =   315
         ItemData        =   "frmStudGeneral.frx":0000
         Left            =   1200
         List            =   "frmStudGeneral.frx":0002
         TabIndex        =   20
         Text            =   "<<--SELECT-->>"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Cmdgo 
         Caption         =   "&Go"
         Height          =   375
         Left            =   5280
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtsuid 
         DataField       =   "student_id"
         DataSource      =   "Adodc1"
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
         Left            =   5280
         TabIndex        =   18
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   25
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label6 
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
         Left            =   675
         TabIndex        =   24
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   7200
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
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
         Left            =   120
         MouseIcon       =   "frmStudGeneral.frx":0004
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdStdGenFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Full record  and  Edit "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   120
         MouseIcon       =   "frmStudGeneral.frx":0156
         MousePointer    =   99  'Custom
         Picture         =   "frmStudGeneral.frx":02A8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student General Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   6975
      Begin VB.TextBox txtstd 
         Height          =   375
         Left            =   1380
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtdiv 
         Height          =   375
         Left            =   3060
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtln 
         Height          =   375
         Left            =   1380
         TabIndex        =   7
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtmn 
         Height          =   375
         Left            =   1380
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtfn 
         Height          =   375
         Left            =   1380
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Image pcbox 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   4380
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Photograph"
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
         Height          =   375
         Left            =   4980
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Last name :"
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
         Left            =   315
         TabIndex        =   12
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle name :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "First name :"
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
         Left            =   300
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DIV :"
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
         Left            =   2595
         TabIndex        =   4
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
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
         Left            =   720
         TabIndex        =   3
         Top             =   2280
         Width           =   570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexMember 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   12640511
      BackColorSel    =   255
      BackColorBkg    =   16777215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student General Info :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmStudGeneral.frx":0929
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmStudGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub ComStd_Click()
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


Private Sub ComDiv_Click()
Call connect
Dim sql As String
If rs_stugrid.State = adStateOpen Then rs_stugrid.Close
rs_stugrid.CursorLocation = adUseClient
sql = "select * from student_mstr where Std= '" & ComStd.Text & "' and Div= '" & ComDiv.Text & "'"
rs_stugrid.Open sql, con, adOpenKeyset, adLockOptimistic
If rs_stugrid.RecordCount <= 0 Then
MsgBox "no record found"
ComStd.Text = ""
ComDiv.Text = ""
txtsuid.Text = ""
txtfn.Text = ""
txtmn.Text = ""
txtln.Text = ""
txtstd.Text = ""
txtdiv.Text = ""
pcbox.Picture = Nothing

End If
   With FlexMember
      FlexMember.Cols = rs_stugrid.Fields.count + 1
      FlexMember.ColWidth(0) = 0
        For c = 0 To rs_stugrid.Fields.count - 1
          FlexMember.TextMatrix(0, c + 1) = rs_stugrid(c).Name
        Next
      FlexMember.Rows = rs_stugrid.RecordCount + 1
        For r = 1 To rs_stugrid.RecordCount
           For c = 0 To rs_stugrid.Fields.count - 1
             FlexMember.TextMatrix(r, c + 1) = IIf(IsNull(rs_stugrid(c).Value), "{Null}", rs_stugrid(c).Value)
           Next c
          rs_stugrid.MoveNext
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

End Sub


Private Sub cmdClassfindOk_Click()
Unload Me
End Sub

Private Sub Cmdgo_Click()

txtfn.Text = ""
txtmn.Text = ""
txtln.Text = ""
txtstd.Text = ""
txtdiv.Text = ""


On Error Resume Next
Call connect

With rs_find
If .State = adStateOpen Then .Close
.Open " select * from student_mstr where student_id = " & Val(txtsuid.Text) & "", con, adOpenDynamic, adLockPessimistic
If .RecordCount <= 0 Then
MsgBox "no record found"
.Close
txtsuid.Text = ""
Exit Sub
End If
End With

With rs_find
Dim c As Double
c = Val(txtsuid.Text)
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where student_id = " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
txtfn.Text = .Fields("First_name").Value
txtmn.Text = .Fields("Middle_name").Value
txtln.Text = .Fields("Last_name").Value
txtstd.Text = .Fields("Std").Value
txtdiv.Text = .Fields("Div").Value
Dim a As String
a = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(a)
.MoveNext
Loop
.Close
End With



End Sub

Private Sub cmdStdGenFull_Click()
Me.Hide
frmStudFull.Show
End Sub



Public Function fillgrid()
Call connect
Dim sql As String
rs_stugrid.CursorLocation = adUseClient
sql = "select * from student_mstr"
rs_stugrid.Open sql, con, adOpenKeyset, adLockOptimistic
   With FlexMember
      FlexMember.Cols = rs_stugrid.Fields.count + 1
      FlexMember.ColWidth(0) = 0
        For c = 0 To rs_stugrid.Fields.count - 1
          FlexMember.TextMatrix(0, c + 1) = rs_stugrid(c).Name
        Next
      FlexMember.Rows = rs_stugrid.RecordCount + 1
        For r = 1 To rs_stugrid.RecordCount
           For c = 0 To rs_stugrid.Fields.count - 1
             FlexMember.TextMatrix(r, c + 1) = IIf(IsNull(rs_stugrid(c).Value), "{Null}", rs_stugrid(c).Value)
           Next c
          rs_stugrid.MoveNext
        Next r
   End With

End Function

Private Sub FlexMember_Click()
On Error Resume Next
txtsuid.Text = ""
txtfn.Text = ""
txtmn.Text = ""
txtln.Text = ""
txtstd.Text = ""
txtdiv.Text = ""
pcbox.Picture = Nothing

Call connect
rs_stugrid.Open "select * from student_mstr", con, adOpenDynamic, adLockOptimistic
rs_stugrid.MoveFirst
rs_stugrid.Move FlexMember.Row - 1
txtsuid.Text = FlexMember.TextMatrix(FlexMember.Row, 1)
txtfn.Text = FlexMember.TextMatrix(FlexMember.Row, 2)
txtmn.Text = FlexMember.TextMatrix(FlexMember.Row, 3)
txtln.Text = FlexMember.TextMatrix(FlexMember.Row, 4)
txtstd.Text = FlexMember.TextMatrix(FlexMember.Row, 5)
txtdiv.Text = FlexMember.TextMatrix(FlexMember.Row, 6)
a = App.Path & FlexMember.TextMatrix(FlexMember.Row, 42)
pcbox.Picture = LoadPicture(a)
End Sub

Private Sub Form_Load()
Call connect
fillgrid

With rs_find
If .State = adStateOpen Then .Close
.Open "select distinct Std from class_mstr order by Std", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
ComStd.AddItem .Fields("Std")
.MoveNext
Loop
.Close
End With
End Sub
