VERSION 5.00
Begin VB.Form frmUserAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User Registration"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   4200
      Width           =   5295
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Create"
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
         Left            =   1080
         MouseIcon       =   "frmUserAdd.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdcan 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Cancel"
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
         Left            =   2640
         MouseIcon       =   "frmUserAdd.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5295
      Begin VB.TextBox txtname 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtrepass 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox comstffid 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmUserAdd.frx":02A4
         Left            =   2520
         List            =   "frmUserAdd.frx":02A6
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtun 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox comAcctyp 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmUserAdd.frx":02A8
         Left            =   2520
         List            =   "frmUserAdd.frx":02B2
         TabIndex        =   2
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2250
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "New User Registration"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frmUserAdd.frx":02D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmUserAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCan_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
Call connect
If txtun.Text = "" Or txtPass.Text = "" Or txtrepass.Text = "" Or txtname.Text = "" Or comAcctyp.Text = "" Or comAcctyp.Text = "<<--SELECT-->>" Then
MsgBox "all fill all fields"
Exit Sub
End If


If txtPass.Text <> txtrepass.Text Then
MsgBox " please confirm password"
txtrepass.Text = ""
Exit Sub
End If



With rs_find
If .State = adStateOpen Then .Close
.Open " select * from user_mstr where staff_id= " & Val(comstffid.Text) & "", con, adOpenDynamic, adLockOptimistic
If .RecordCount >= 1 Then
.Close
MsgBox "current user already present"
comstffid.Text = ""
Exit Sub
End If
.Close
End With


Dim rep
rep = MsgBox("Are you sure you wanna create new User ?", vbYesNo)

If rep = vbYes Then
    With rs_find
        If .State = adStateOpen Then .Close
        .Open "select * from user_mstr", con, adOpenDynamic, adLockPessimistic
        .AddNew
        .Fields("user_id") = txtun.Text
        .Fields("pass") = txtPass.Text
        .Fields("staff_id").Value = Val(comstffid.Text)
        .Fields("Acct_typ") = comAcctyp.Text
        .Update
        .Close
        MsgBox "User Successfully Created !"
    End With
    Unload Me
Else
    Exit Sub
End If
End Sub



Private Sub comstffid_Click()
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr where staff_id= " & Val(comstffid.Text) & "", con, adOpenDynamic, adLockOptimistic
Do Until .EOF
txtname.Text = .Fields("fname") & " " & .Fields("mname") & " " & .Fields("lname")
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Form_Load()
Call CenterForm(frmUserAdd)
Call connect
With rs_find
If .State = adStateOpen Then .Close
.Open "select staff_id from staff_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
comstffid.AddItem .Fields("staff_id").Value
.MoveNext
Loop
.Close
End With
End Sub
