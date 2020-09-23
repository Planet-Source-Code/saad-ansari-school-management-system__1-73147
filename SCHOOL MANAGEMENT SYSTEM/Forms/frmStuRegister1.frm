VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsturegister1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Enrollment form I"
   ClientHeight    =   9060
   ClientLeft      =   3075
   ClientTop       =   1335
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComGen 
      Height          =   315
      ItemData        =   "frmStuRegister1.frx":0000
      Left            =   8160
      List            =   "frmStuRegister1.frx":000A
      TabIndex        =   45
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Texfname 
      Height          =   375
      Left            =   3960
      TabIndex        =   43
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Texmname 
      Height          =   375
      Left            =   3960
      TabIndex        =   41
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Texage 
      Height          =   375
      Left            =   8160
      TabIndex        =   39
      Top             =   6960
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   600
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Texnaton 
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox TexPOB 
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox Texmotname 
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Texaddress 
      Height          =   855
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Texcity 
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Texstate 
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox Texpin 
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   6840
      Width           =   2535
   End
   Begin VB.TextBox Texphone 
      Height          =   375
      Left            =   3960
      MaxLength       =   14
      TabIndex        =   13
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox Texmobile 
      Height          =   375
      Left            =   3960
      MaxLength       =   14
      TabIndex        =   12
      Top             =   7800
      Width           =   2535
   End
   Begin VB.TextBox Texlname 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Texgarname 
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Texfatname 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Texemail 
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdclr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdcls 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox Texstu_id 
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtdob 
      Height          =   375
      Left            =   8160
      TabIndex        =   46
      Top             =   6000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   40251
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmStuRegister1.frx":001C
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6840
      TabIndex        =   47
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label19 
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
      Left            =   2745
      TabIndex        =   44
      Top             =   2160
      Width           =   975
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
      Left            =   2565
      TabIndex        =   42
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identity"
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
      Height          =   255
      Left            =   7080
      TabIndex        =   40
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
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
      Left            =   7485
      TabIndex        =   38
      Top             =   6960
      Width           =   420
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   37
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Image pcbox 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality :"
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
      Left            =   6975
      TabIndex        =   36
      Top             =   6480
      Width           =   930
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Place :"
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
      Left            =   6945
      TabIndex        =   35
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :"
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
      TabIndex        =   34
      Top             =   5040
      Width           =   705
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DOB :"
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
      Left            =   7485
      TabIndex        =   33
      Top             =   6000
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
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
      Left            =   3315
      TabIndex        =   30
      Top             =   6000
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
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
      Left            =   3210
      TabIndex        =   29
      Top             =   6480
      Width           =   510
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
      Left            =   2760
      TabIndex        =   28
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent address :*"
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
      Left            =   1980
      TabIndex        =   27
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode :"
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
      Left            =   2970
      TabIndex        =   26
      Top             =   6960
      Width           =   750
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number :"
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
      Left            =   2265
      TabIndex        =   25
      Top             =   7440
      Width           =   1425
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No. :"
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
      Left            =   2775
      TabIndex        =   24
      Top             =   7920
      Width           =   945
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Father name :"
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
      Left            =   2595
      TabIndex        =   23
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mother name :"
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
      Left            =   2520
      TabIndex        =   22
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian Name (if any) :"
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
      Left            =   1800
      TabIndex        =   21
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
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
      Left            =   6840
      TabIndex        =   20
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address :"
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
      Left            =   2430
      TabIndex        =   19
      Top             =   8400
      Width           =   1290
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1/2.  Click next to proceed and Close to Cancel the Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Label Label17 
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
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmStuRegister1.frx":00A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   5295
      Left            =   6720
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   7335
      Left            =   1560
      Top             =   1680
      Width           =   9615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   7335
      Left            =   1680
      Top             =   1800
      Width           =   9615
   End
End
Attribute VB_Name = "frmsturegister1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public str As Variant
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim no As Integer




Private Sub Command3_Click()
If Texfname.Text = "" Or Texmname.Text = "" Or Texlname.Text = "" Or _
    Texmotname.Text = "" Or Texfatname.Text = "" Or Texgarname.Text = "" Or _
    Texaddress.Text = "" Or Texcity.Text = "" Or Texstate.Text = "" Or _
    Texpin.Text = "" Or Texphone.Text = "" Or Texmobile.Text = "" Or _
    Texemail.Text = "" Or ComGen.Text = "" Or TexPOB.Text = "" Or _
    Texnaton.Text = "" Or Texage.Text = "" Then
    MsgBox "Please Fill  all the entries", vbInformation + vbOKOnly, "FIELD EMPTY"
Else
Me.Hide
frmsturegister2.Show
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
 On Error Resume Next
    cdb.Filter = "All Picture Files|*.jpg;*.gif;*.bmp;*.wmf;*.ico|JPEG Images(*.jpg)|*.jpg|Bitmap Images (*.bmp)|*.bmp|Word Meta Files (*.wmf)|*.wmf|GIF Images (*.gif)|*.gif"
    cdb.ShowOpen
    If cdb.FileName <> "" Then
        pcbox.Picture = LoadPicture(cdb.FileName)
        pic_name = cdb.FileName
        pic_ext = Right(cdb.FileTitle, 4)
        pic_changed = True
    End If
End Sub

Private Sub cmdclr_Click()
 On Error Resume Next
    Set pcbox.Picture = Nothing
    pic_name = ""
    pic_changed = True

End Sub

Private Sub cmdcls_Click()
Unload Me
End Sub

Private Sub cmdnext_Click()
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
Me.Hide
frmsturegister2.Show
End Sub

Private Sub dtdob_Click()
Dim a, b, c As Integer
a = Format(dtdob.Value, "yyyy")
b = Format(Date, "yyyy")
c = b - a
Texage.Text = c
End Sub

Private Sub Form_Load()
On Error Resume Next
Call connect
con.Refresh
With rs_find
.Open "select * from student_mstr", con, adOpenDynamic, adLockPessimistic
.MoveLast
If IsNull(.Fields("student_id").Value) Then
Texstu_id.Text = 1
Else
no = .Fields("student_id") + 1
Texstu_id.Text = no
End If
.Close
End With

End Sub

Private Sub Texage_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texfatname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub Texfname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub Texlname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub Texmname_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub Texmobile_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texphone_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texpin_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
