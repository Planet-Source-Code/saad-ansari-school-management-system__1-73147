VERSION 5.00
Begin VB.Form frmedtDep 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edit Department "
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   6735
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton cmdedtsav 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         MouseIcon       =   "frmedtDep.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Edit Record"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MouseIcon       =   "frmedtDep.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Delete Record"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         MouseIcon       =   "frmedtDep.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Close"
         Height          =   375
         Left            =   4200
         MouseIcon       =   "frmedtDep.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox comdepid 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2280
         MouseIcon       =   "frmedtDep.frx":0548
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox Texdn 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Texdes 
         BackColor       =   &H00C0FFFF&
         Height          =   1695
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department ID :"
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
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
         Left            =   840
         TabIndex        =   8
         Top             =   1680
         Width           =   1170
      End
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Department :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -240
      Picture         =   "frmedtDep.frx":0E12
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmedtDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmddel_Click()
Dim strsql
Dim rep
rep = MsgBox("Are you sure you want to permanently delete this record ?", vbQuestion + vbYesNo, "Confirm Record Delete")
If rep = vbYes Then
    Call connect
    strsql = "delete * from department where department_id = " & Val(comdepid.Text) & ""
    con.Execute strsql
    MsgBox "Record Successfully Deleted"
    comdepid.RemoveItem (comdepid.ListIndex)
    Texdes = ""
    Texdn = ""
Else
    Exit Sub
End If
End Sub

Private Sub cmdEdit_Click()
cmdedtsav.Enabled = True
Texdes.Locked = False
Texdn.Locked = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdedtsav_Click()
If Texdn.Text = "" Then
MsgBox "Please provide a department name !!"
Exit Sub
End If
Call connect
Dim strsql
strsql = "UPDATE  department set department_name='" & Texdn.Text & _
        "',description='" & Texdes.Text & "' where department_id=" & Val(comdepid.Text)
con.Execute strsql
MsgBox "Record Successfully Updated"
Texdes.Locked = True
Texdn.Locked = True
cmdEdit.Enabled = True
cmdedtsav.Enabled = False
End Sub

Private Sub Comdepid_Click()
Call connect
cmdDel.Enabled = True
cmdEdit.Enabled = True
With rs_find
Dim c As Double
c = Val(comdepid.Text)
If .State = adStateOpen Then .Close
.Open "select * from department where department_id=" & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texdn.Text = .Fields("department_name")
Texdes.Text = .Fields("description")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from department", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comdepid.AddItem .Fields("department_id")
.MoveNext
Loop
.Close
End With
End Sub
