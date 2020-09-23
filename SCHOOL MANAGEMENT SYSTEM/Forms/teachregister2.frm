VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form teachregister2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Staff Registration"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Comtype 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "teachregister2.frx":0000
      Left            =   7440
      List            =   "teachregister2.frx":000A
      TabIndex        =   43
      Top             =   5160
      Width           =   1815
   End
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Save"
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Texrem 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Left            =   8760
      TabIndex        =   37
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Texpre 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Left            =   8760
      TabIndex        =   36
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Texnoyear 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Left            =   8760
      TabIndex        =   35
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Comexp 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "teachregister2.frx":0026
      Left            =   8880
      List            =   "teachregister2.frx":0030
      TabIndex        =   29
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Texsub4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   27
      Top             =   7560
      Width           =   1695
   End
   Begin VB.ComboBox Comjob 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "teachregister2.frx":004A
      Left            =   3120
      List            =   "teachregister2.frx":0054
      TabIndex        =   25
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdcls 
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8520
      Width           =   1095
   End
   Begin VB.ComboBox Comdepid 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3120
      TabIndex        =   23
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Texsub3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   22
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Texsub2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   21
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Texsub1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Texsal 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Texenroll 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Texqual 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Texdepname 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Texsumqual 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Texpost 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5280
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPendate 
      Height          =   375
      Left            =   3120
      TabIndex        =   41
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58982401
      CurrentDate     =   40251
   End
   Begin MSComCtl2.DTPicker DTPCT 
      Height          =   375
      Left            =   3120
      TabIndex        =   42
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58982401
      CurrentDate     =   40251
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"teachregister2.frx":006E
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6240
      TabIndex        =   44
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
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
      Left            =   6720
      TabIndex        =   39
      Top             =   5160
      Width           =   585
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   1935
      Left            =   6720
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Employer :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   34
      Top             =   4320
      Width           =   1965
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   33
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of years :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   32
      Top             =   3240
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If Experienced :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6840
      TabIndex        =   31
      Top             =   2640
      Width           =   1560
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Experience :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   30
      Top             =   2040
      Width           =   2130
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 4 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7380
      TabIndex        =   28
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUBJECT ASSIGNED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2415
      Left            =   6720
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 3 : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7380
      TabIndex        =   19
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 2 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7380
      TabIndex        =   18
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject 1 :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7380
      TabIndex        =   17
      Top             =   6120
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2250
      TabIndex        =   15
      Top             =   7200
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1305
      TabIndex        =   14
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Termination :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   810
      TabIndex        =   13
      Top             =   7800
      Width           =   2190
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Registration : "
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "teachregister2.frx":00F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Status :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1875
      TabIndex        =   8
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1650
      TabIndex        =   7
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1110
      TabIndex        =   5
      Top             =   2520
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment Year :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1305
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary of Qualification :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   405
      TabIndex        =   3
      Top             =   4320
      Width           =   2595
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2460
      TabIndex        =   2
      Top             =   5400
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6735
      Left            =   240
      Top             =   1560
      Width           =   10935
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6735
      Left            =   360
      Top             =   1680
      Width           =   10935
   End
End
Attribute VB_Name = "teachregister2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pic_name As String, pic_ext As String, pic_changed As Boolean

Private Sub cmd_Bk_Click()
teachregister.Show
Unload Me
End Sub

Private Sub cmdcls_Click()
Unload teachregister
Unload Me
End Sub

Private Sub cmdSave_Click()

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

With rs_staff
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr", con, adOpenDynamic, adLockOptimistic
.AddNew
.Fields("staff_id").Value = teachregister.Texid.Text
.Fields("fname") = teachregister.Texfname.Text
.Fields("mname") = teachregister.Texmname.Text
.Fields("lname") = teachregister.Texlname.Text
.Fields("dob") = teachregister.DTPdob.Value
.Fields("sex") = teachregister.Comsex.Text
.Fields("address") = teachregister.Texadd.Text
.Fields("city") = teachregister.Texcity.Text
.Fields("state") = teachregister.Texstate.Text
.Fields("pincode") = teachregister.Texpin.Text
.Fields("con_no") = teachregister.Texcn.Text
.Fields("mob_no") = teachregister.Texmo.Text
.Fields("email") = teachregister.Texemail.Text
.Fields("nationality") = teachregister.Texnat.Text
.Fields("birthplace") = teachregister.Texbp.Text
.Fields("age") = teachregister.Texage.Text
.Fields("religion") = teachregister.Texrel.Text
.Fields("caste") = teachregister.Texcast.Text
.Fields("civilstatus") = teachregister.Texcivil.Text
.Fields("type") = Comtype.Text
.Fields("department_id") = comdepid.Text
.Fields("department_name") = Texdepname.Text
.Fields("qualification") = Texqual.Text
.Fields("summaryofqual") = Texsumqual.Text
.Fields("post") = Texpost.Text
.Fields("enroll_yr") = Texenroll.Text
.Fields("jobstatus") = Comjob.Text
.Fields("enrollDate") = DTPendate.Value
.Fields("salary") = texsal.Text
.Fields("cont_term") = DTPCT.Value
.Fields("pre_exp") = Comexp.Text
.Fields("no_of_yr") = Texnoyear.Text
.Fields("remark") = Texrem.Text
.Fields("pre_employr") = Texpre.Text
.Fields("Subject1") = Texsub1.Text
.Fields("Subject2") = Texsub2.Text
.Fields("Subject3") = Texsub3.Text
.Fields("Subject4") = Texsub4.Text
If pic_name <> "" Then
                    FileCopy pic_name, App.Path & "\Miscellaneous\STAFF_IMAGE\" & .Fields("staff_id") & pic_ext
                    .Fields("picture") = "\Miscellaneous\STAFF_IMAGE\" & .Fields("staff_id") & pic_ext

End If

.Update
.Close
MsgBox "Record Successfully Entered", vbInformation + vbOKOnly, "NEW RECORD ENTRY"
End With
Dim rep
   rep = MsgBox("Do you wish to add a new STAFF ?", vbInformation + vbYesNo, "CREATE NEW")
   If rep = vbYes Then
       Unload teachregister
       Unload Me
       teachregister.Show
       
   Else
       Unload teachregister
       Unload Me
   End If




End Sub
Private Sub Comtype_Click()
If Comtype.Text = "NON-TEACHING" Then
teachregister2.Texsub1.Text = "NiL"
teachregister2.Texsub2.Text = "NiL"
teachregister2.Texsub3.Text = "NiL"
teachregister2.Texsub4.Text = "NiL"

teachregister2.Texsub1.Enabled = False
teachregister2.Texsub2.Enabled = False
teachregister2.Texsub3.Enabled = False
teachregister2.Texsub4.Enabled = False
Else
teachregister2.Texsub1.Text = ""
teachregister2.Texsub2.Text = ""
teachregister2.Texsub3.Text = ""
teachregister2.Texsub4.Text = ""
teachregister2.Texsub1.Enabled = True
teachregister2.Texsub2.Enabled = True
teachregister2.Texsub3.Enabled = True
teachregister2.Texsub4.Enabled = True
End If
End Sub

Private Sub Comdepid_Click()
With rs_find
If .State = adStateOpen Then .Close
.Open "select department_name from department where department_id= " & Val(comdepid.Text) & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texdepname.Text = .Fields("department_name")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Comexp_Click()
If Comexp.Text = "EXPERIENCED" Then
Texnoyear.Enabled = True
Texrem.Enabled = True
Texpre.Enabled = True
Else
Texnoyear.Text = "NiL"
Texrem.Text = "NiL"
Texpre.Text = "NiL"

Texnoyear.Enabled = False
Texrem.Enabled = False
Texpre.Enabled = False
End If
End Sub

Private Sub Form_Load()
If teachregister.cdb.FileName <> "" Then
       teachregister.pcbox.Picture = LoadPicture(teachregister.cdb.FileName)
        pic_name = teachregister.cdb.FileName
        pic_ext = Right(teachregister.cdb.FileTitle, 4)
        pic_changed = True
    End If
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

Private Sub Texenroll_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texnoyear_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texsal_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
