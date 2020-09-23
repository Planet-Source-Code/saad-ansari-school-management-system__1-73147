VERSION 5.00
Begin VB.Form editclass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edit Class Information"
   ClientHeight    =   7365
   ClientLeft      =   1935
   ClientTop       =   810
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7960.354
   ScaleMode       =   0  'User
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MouseIcon       =   "editclass.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MouseIcon       =   "editclass.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Edit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MouseIcon       =   "editclass.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox TexPerstu 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   20
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Texclassrec 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Texstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Texdiv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox Comsch 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "editclass.frx":03F6
      Left            =   3300
      List            =   "editclass.frx":0400
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox sub8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox sub7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox sub6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox sub5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox sub4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox sub3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox sub2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox sub1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox Comacdbehav 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "editclass.frx":0418
      Left            =   3300
      List            =   "editclass.frx":0431
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox sub9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox sub10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox sub11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox sub12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.ComboBox Comclid 
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permitted Student Strength :"
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
      Left            =   480
      TabIndex        =   38
      Top             =   2880
      Width           =   2715
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class requirement :"
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
      Left            =   1350
      TabIndex        =   37
      Top             =   4320
      Width           =   1845
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acemedic Behaviour :"
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
      Left            =   1080
      TabIndex        =   36
      Top             =   3840
      Width           =   2100
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class shedule :"
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
      Left            =   1770
      TabIndex        =   35
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label Label6 
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
      Height          =   375
      Left            =   2700
      TabIndex        =   34
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   2700
      TabIndex        =   33
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   32
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   31
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   23
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   22
      Top             =   6000
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   21
      Top             =   6480
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class ID :"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Class Information :"
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
      Height          =   675
      Left            =   0
      Picture         =   "editclass.frx":0471
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6015
      Left            =   240
      Top             =   960
      Width           =   8775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   6015
      Left            =   360
      Top             =   1080
      Width           =   8775
   End
End
Attribute VB_Name = "editclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
TexPerstu.Locked = False
Comsch.Locked = False
Comacdbehav.Locked = False
Texclassrec.Locked = False
sub1.Locked = False
sub2.Locked = False
sub3.Locked = False
sub4.Locked = False
sub5.Locked = False
sub6.Locked = False
sub7.Locked = False
sub8.Locked = False
sub9.Locked = False
sub10.Locked = False
sub11.Locked = False
sub12.Locked = False
cmdsave.Enabled = True
cmdEdit.Enabled = False
Texstd.Enabled = False
Texdiv.Enabled = False
End Sub

Private Sub cmdSave_Click()
Call connect
 If sub1.Text = "" And sub1.Text = "" And sub2.Text = "" And sub3.Text = "" And sub4.Text = "" And _
sub5.Text = "" And sub6.Text = "" And sub7.Text = "" And sub8.Text = "" And sub9.Text = "" And _
sub10.Text = "" And sub11.Text = "" And sub12.Text = "" Then
MsgBox "please provide atleast one subject"
Exit Sub
End If

Dim strsql, syrsql1, strsql2
strsql = "UPDATE  class_mstr set Std='" & Texstd.Text & _
"',Div='" & Texdiv.Text & "',Student_Strength='" & TexPerstu.Text & _
"',Class_shedule='" & Comsch.Text & _
"',Acemedic_Behaviour='" & Comacdbehav.Text & "',Class_requirenment='" & Texclassrec.Text & _
"',Subject1='" & sub1.Text & "',Subject2='" & sub2.Text & _
"',Subject3='" & sub3.Text & "',Subject4='" & sub4.Text & _
"',Subject6='" & sub6.Text & "',Subject5='" & sub5.Text & _
"',Subject7='" & sub7.Text & "',Subject8='" & sub8.Text & _
"',Subject10='" & sub10.Text & "',Subject9='" & sub9.Text & _
"',Subject11='" & sub11.Text & "',Subject12='" & sub12.Text & _
"' where Class_ID =" & Val(Comclid.Text)

con.Execute strsql
MsgBox "data updated"

TexPerstu.Locked = True
Comsch.Locked = True
Comacdbehav.Locked = True
Texclassrec.Locked = True
sub1.Locked = True
sub2.Locked = True
sub3.Locked = True
sub4.Locked = True
sub5.Locked = True
sub6.Locked = True
sub7.Locked = True
sub8.Locked = True
sub9.Locked = True
sub10.Locked = True
sub11.Locked = True
sub12.Locked = True
cmdsave.Enabled = False
cmdEdit.Enabled = True
Texstd.Enabled = True
Texdiv.Enabled = True

End Sub

Private Sub Comclid_Click()
On Error Resume Next
Texstd.Text = ""
Texdiv.Text = ""
TexPerstu.Text = ""
Texteach.Text = ""
Texstffid.Text = ""
Comsch.Text = ""
Comacdbehav.Text = ""
Texclassrec.Text = ""
sub1.Text = ""
sub2.Text = ""
sub3.Text = ""
sub4.Text = ""
sub5.Text = ""
sub6.Text = ""
sub7.Text = ""
sub8.Text = ""
sub9.Text = ""
sub10.Text = ""
sub11.Text = ""
sub12.Text = ""

Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from class_mstr where Class_ID = " & Val(Comclid.Text) & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texstd.Text = .Fields("Std")
Texdiv.Text = .Fields("Div")
TexPerstu.Text = .Fields("Student_Strength")
Texteach.Text = .Fields("Class_teacher")
Comsch.Text = .Fields("Class_shedule")
Texstffid.Text = .Fields("staff_id")
Comacdbehav.Text = .Fields("Acemedic_Behaviour")
Texclassrec.Text = .Fields("Class_requirenment")
sub1.Text = .Fields("Subject1")
sub2.Text = .Fields("Subject2")
sub3.Text = .Fields("Subject3")
sub4.Text = .Fields("Subject4")
sub5.Text = .Fields("Subject5")
sub6.Text = .Fields("Subject6")
sub7.Text = .Fields("Subject7")
sub8.Text = .Fields("Subject8")
sub9.Text = .Fields("Subject9")
sub10.Text = .Fields("Subject10")
sub11.Text = .Fields("Subject11")
sub12.Text = .Fields("Subject12")
.MoveNext
Loop
.Close
End With
cmdEdit.Enabled = True
End Sub




Private Sub Form_Load()
Call connect

With rs_find
If .State = adStateOpen Then .Close
.Open "select * from class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Comclid.AddItem .Fields("Class_ID")
.MoveNext
Loop
.Close
End With
End Sub

