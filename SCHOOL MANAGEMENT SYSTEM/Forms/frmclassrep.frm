VERSION 5.00
Begin VB.Form frmclassrep 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Class Report"
   ClientHeight    =   2715
   ClientLeft      =   1710
   ClientTop       =   3090
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   5520
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox ComDiv 
         Height          =   315
         ItemData        =   "frmclassrep.frx":0000
         Left            =   3360
         List            =   "frmclassrep.frx":0002
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox ComStd 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1695
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Std :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.CommandButton cmscls 
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3960
      MouseIcon       =   "frmclassrep.frx":0004
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Show"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmclassrep.frx":0156
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Report"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmclassrep.frx":02A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmclassrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdshow_Click()
Call connect
With rs_find
Dim count As Integer
count = 0
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where Std='" & comStd.Text & "' and Div='" & comDiv.Text & "'", con, adOpenStatic, adLockPessimistic
If .RecordCount <= 0 Then
            MsgBox "No Entry For This calss !", vbExclamation, Me.Caption
            Exit Sub
End If
Set classreport.DataSource = rs_find
classreport.Sections("section4").Controls("class").Caption = comStd.Text & " " & comDiv.Text
.MoveFirst
If .EOF = False Then
    Do Until .EOF
    count = count + 1
    .MoveNext
    Loop
End If
classreport.Sections("section4").Controls("lbltotal_max_marks").Caption = count
classreport.Show

End With
Unload Me
End Sub

Private Sub cmscls_Click()
Unload Me
End Sub

Private Sub ComStd_Click()
comDiv.Clear
Call connect
With rs_class
If .State = adStateOpen Then .Close
.Open "select distinct Div from class_mstr where Std = '" & comStd.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comDiv.AddItem .Fields("Div")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub Form_Load()
Call connect
With rs_class
If .State = adStateOpen Then .Close
.Open "SELECT distinct Std FROM class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comStd.AddItem .Fields("Std")
.MoveNext
Loop
.Close
End With

End Sub
