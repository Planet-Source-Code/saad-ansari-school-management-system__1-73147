VERSION 5.00
Begin VB.Form Findclass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Find Class"
   ClientHeight    =   6780
   ClientLeft      =   1335
   ClientTop       =   2235
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   8475
   WindowState     =   2  'Maximized
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
      Left            =   7200
      MouseIcon       =   "Findclass.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   975
   End
   Begin VB.ComboBox Comcid 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   8175
      Begin VB.TextBox Texstrength 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Texnos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Texab 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox Texstd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Texdiv 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Texto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Texaf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Texcf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Texef 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Texac 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Textf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Texgf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Permited"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4320
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of student :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4215
         TabIndex        =   25
         Top             =   2520
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acemedic Behaviour :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4080
         TabIndex        =   24
         Top             =   3000
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Div :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5265
         TabIndex        =   23
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Std :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5265
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label12 
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1770
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label Label10 
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
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2010
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   1845
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   11
         Top             =   3720
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
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   1860
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Info :"
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
      TabIndex        =   17
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Class ID:-"
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
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   0
      Picture         =   "Findclass.frx":030A
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   8475
   End
End
Attribute VB_Name = "Findclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClassfindOk_Click()
Unload Me
End Sub

Private Sub Comcid_Click()
'On Error Resume Next
Call connect
Textf.Text = ""
Texgf.Text = ""
Texac.Text = ""
Texef.Text = ""
Texcf.Text = ""
Texaf.Text = ""
Texto.Text = ""
Texstd.Text = ""
Texdiv.Text = ""
Texnos.Text = ""
Texstrength.Text = ""
Texab.Text = ""

With rs_find
If .State = adStateOpen Then .Close
.Open "select * from class_mstr where Class_ID = " & Val(Comcid.Text) & "", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
Texstd.Text = .Fields("Std").Value
Texdiv.Text = .Fields("Div").Value
Texab.Text = .Fields("Acemedic_Behaviour").Value
Texstrength.Text = .Fields("Student_Strength")
.MoveNext
Loop
.Close
End With


With rs_find
If .State = adStateOpen Then .Close
.Open "select student_id from student_mstr where Std='" & Texstd.Text & "' and Div='" & Texdiv.Text & "'", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Dim d As Integer
d = d + 1
.MoveNext
Loop
.Close
End With
Texnos.Text = d






With rs_find
If .State = adStateOpen Then .Close
.Open "select Tution_Fees,General_Fund,Annual_Charges,Examination_Fee,Computer_Fee,Admission_Fee,Total from fees_stru where fees_stru.Class_ID =" & Val(Comcid.Text) & "", con, adOpenDynamic, adLockPessimistic
If .RecordCount <= 0 Then
.Close
MsgBox "no fees structure for this class"
Exit Sub
End If

Do Until .EOF
Textf.Text = .Fields("Tution_Fees").Value
Texgf.Text = .Fields("General_Fund").Value
Texac.Text = .Fields("Annual_Charges").Value
Texef.Text = .Fields("Examination_Fee").Value
Texcf.Text = .Fields("Computer_Fee").Value
Texaf.Text = .Fields("Admission_Fee").Value
Texto.Text = .Fields("Total").Value
.MoveNext
Loop
.Close
End With


End Sub

Private Sub Form_Load()
Call CenterForm(Me)
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select Class_ID from class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Comcid.AddItem .Fields("Class_ID").Value
.MoveNext
Loop
.Close
End With
End Sub

