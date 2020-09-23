VERSION 5.00
Begin VB.Form Addclass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Class Information"
   ClientHeight    =   7845
   ClientLeft      =   1080
   ClientTop       =   1725
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   9855
      Begin VB.TextBox Texclassid 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdSav 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Save"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         MouseIcon       =   "Classinfo.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmdclos 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         MouseIcon       =   "Classinfo.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox sub12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox sub11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox sub10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CommandButton Cmdadd 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add"
         Height          =   375
         Left            =   1680
         MouseIcon       =   "Classinfo.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox sub9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ComboBox Comacdbehav 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Classinfo.frx":03F6
         Left            =   2940
         List            =   "Classinfo.frx":040F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox sub1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox sub2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox sub3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox sub4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox sub5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox sub6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox sub7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox sub8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3840
         Width           =   1695
      End
      Begin VB.ComboBox Comsch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "Classinfo.frx":044F
         Left            =   2940
         List            =   "Classinfo.frx":0459
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Texdiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Texstd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Texclassrec 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox TexPerstu 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         MouseIcon       =   "Classinfo.frx":0471
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6240
         Width           =   1455
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
         Left            =   1680
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
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
         Left            =   6840
         TabIndex        =   43
         Top             =   5880
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
         Left            =   6840
         TabIndex        =   42
         Top             =   5400
         Width           =   990
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
         Left            =   6840
         TabIndex        =   41
         Top             =   4920
         Width           =   990
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
         Left            =   6840
         TabIndex        =   40
         Top             =   4440
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
         Left            =   6840
         TabIndex        =   39
         Top             =   600
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
         Left            =   6840
         TabIndex        =   38
         Top             =   1080
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
         Left            =   6840
         TabIndex        =   37
         Top             =   1560
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
         Left            =   6840
         TabIndex        =   36
         Top             =   2040
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
         Left            =   6840
         TabIndex        =   35
         Top             =   2520
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
         Left            =   6840
         TabIndex        =   34
         Top             =   3000
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
         Left            =   6840
         TabIndex        =   33
         Top             =   3480
         Width           =   975
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
         Left            =   6840
         TabIndex        =   32
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Left            =   2340
         TabIndex        =   31
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
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
         Left            =   2340
         TabIndex        =   30
         Top             =   1560
         Width           =   495
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
         Left            =   1410
         TabIndex        =   29
         Top             =   2520
         Width           =   1425
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
         Left            =   720
         TabIndex        =   28
         Top             =   2880
         Width           =   2100
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
         Left            =   990
         TabIndex        =   27
         Top             =   3360
         Width           =   1845
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
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subjects :"
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
         Left            =   5640
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Class Information :"
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
      TabIndex        =   22
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   0
      Picture         =   "Classinfo.frx":05C3
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   10100
   End
End
Attribute VB_Name = "Addclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim no As Integer

Private Sub cmdAdd_Click()
cmdSav.Enabled = True
Command4.Enabled = True
cmdclos.Enabled = False
Texclassid.Locked = False
Texstd.Locked = False
Texdiv.Locked = False
TexPerstu.Locked = False
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
Comsch.Locked = False
Comacdbehav.Locked = False
Texstd.Text = ""
Texdiv.Text = ""
TexPerstu.Text = ""
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
cmdSav.Enabled = True
cmdAdd.Enabled = False
Call connect
End Sub

Private Sub cmdclos_Click()
Unload Me
End Sub

Private Sub cmdSav_Click()
'On Error Resume Next

Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from class_mstr where Std = '" & Texstd.Text & "' and Div = '" & Texdiv.Text & "'", con, adOpenStatic, adLockOptimistic
If .RecordCount >= 1 Then
.Close
MsgBox " class allready preasent"
Texstd.Text = ""
Texdiv.Text = ""
Exit Sub
End If
.Close
End With

If Texclassid.Text = "" Or Texstd.Text = "" Or Texdiv.Text = "" Or _
    TexPerstu.Text = "" Or _
    Comsch.Text = "" Or Comacdbehav.Text = "" Or Texclassrec.Text = "" Then
    MsgBox "All fields are Mandatory to filled", vbInformation + vbOKOnly, "Field Empty"
ElseIf sub1.Text = "" And sub2.Text = "" And sub3.Text = "" And sub4.Text = "" And _
        sub5.Text = "" And sub6.Text = "" And sub7.Text = "" And sub8.Text = "" _
         And sub9.Text = "" And sub10.Text = "" And sub11.Text = "" And sub12.Text = "" Then
         
MsgBox "please add  atleast one subject ", vbInformation + vbOKOnly, "Subject Field Empty"
Else
Call connect

With rs_class
If .State = adStateOpen Then .Close
.Open "select * from class_mstr", con, adOpenDynamic, adLockPessimistic
            .AddNew
            .Fields("Class_ID") = Texclassid.Text
            .Fields("Std") = Texstd.Text
            .Fields("Div") = Texdiv.Text
            .Fields("Student_Strength") = TexPerstu.Text
            .Fields("Class_shedule") = Comsch.Text
            .Fields("Acemedic_Behaviour") = Comacdbehav.Text
            .Fields("Class_requirenment") = Texclassrec.Text
            .Fields("Subject1") = sub1.Text
            .Fields("Subject2") = sub2.Text
            .Fields("Subject3") = sub3.Text
            .Fields("Subject4") = sub4.Text
            .Fields("Subject5") = sub5.Text
            .Fields("Subject6") = sub6.Text
            .Fields("Subject7") = sub7.Text
            .Fields("Subject8") = sub8.Text
            .Fields("Subject9") = sub9.Text
            .Fields("Subject10") = sub10.Text
            .Fields("Subject11") = sub11.Text
            .Fields("Subject12") = sub12.Text
            .Update
            .Close
End With
MsgBox "Class Successfully Created", vbInformation + vbOKOnly, "Add Class"
Unload Me
End If
End Sub

Private Sub Command4_Click()
cmdSav.Enabled = False
Command4.Enabled = False
cmdclos.Enabled = True
cmdAdd.Enabled = True
End Sub


Private Sub Form_Load()
On Error Resume Next
Call connect
con.Refresh
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from class_mstr", con, adOpenDynamic, adLockPessimistic
.MoveLast
If IsNull(.Fields("Class_ID").Value) Then
Texclassid.Text = 1
Else
no = .Fields("Class_ID") + 1
Texclassid.Text = no
End If
.Close
End With

End Sub

Private Sub Texdiv_KeyPress(KeyAscii As Integer)
KeyAscii = uppercharacter(KeyAscii)
End Sub

Private Sub TexPerstu_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texstd_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
