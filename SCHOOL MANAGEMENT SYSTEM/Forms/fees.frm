VERSION 5.00
Begin VB.Form feesstu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fees Structure"
   ClientHeight    =   7155
   ClientLeft      =   1065
   ClientTop       =   1710
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Texdiv 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Comfsid 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdfee_sav 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      MouseIcon       =   "fees.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Texstd 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdFeeSturOk 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&close"
      Height          =   375
      Left            =   4200
      MouseIcon       =   "fees.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdfee_mod 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Create New"
      Height          =   375
      Left            =   1320
      MouseIcon       =   "fees.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yearly Fee Structure "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4335
      Left            =   960
      TabIndex        =   15
      Top             =   2160
      Width           =   4815
      Begin VB.TextBox TexGf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TexTf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TexAc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TexEf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TexCf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TexAf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   3  'I-Beam
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox TexfeeTot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   10
         Top             =   3600
         Width           =   1215
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
         Left            =   600
         TabIndex        =   22
         Top             =   3240
         Width           =   1860
      End
      Begin VB.Label Label9 
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
         Left            =   600
         TabIndex        =   21
         Top             =   3720
         Width           =   600
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   20
         Top             =   2760
         Width           =   1845
      End
      Begin VB.Label Label6 
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
         Left            =   600
         TabIndex        =   19
         Top             =   2280
         Width           =   2010
      End
      Begin VB.Label Label3 
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
         Left            =   600
         TabIndex        =   18
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   17
         Top             =   1320
         Width           =   1770
      End
      Begin VB.Label Label8 
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
         Left            =   600
         TabIndex        =   16
         Top             =   840
         Width           =   1605
      End
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
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
      Left            =   1320
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Feed structure:-"
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
      TabIndex        =   14
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -120
      Picture         =   "fees.frx":03F6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Class ID :"
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
      Left            =   990
      TabIndex        =   0
      Top             =   840
      Width           =   1305
   End
End
Attribute VB_Name = "feesstu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdfee_mod_Click()
Textf.Text = ""
Texgf.Text = ""
Texac.Text = ""
Texef.Text = ""
Texcf.Text = ""
Texaf.Text = ""
TexfeeTot.Text = ""
Textf.Enabled = True
Texgf.Enabled = True
Texef.Enabled = True
Texcf.Enabled = True
Texac.Enabled = True
Texaf.Enabled = True
TexfeeTot.Enabled = True

cmdfee_sav.Enabled = True
cmdfee_mod.Enabled = False
End Sub

Private Sub cmdfee_sav_Click()
    Call connect

With rs_find
If .State = adStateOpen Then .Close
.Open "select * from fees_stru where fees_stru.Class_ID = " & Val(Comfsid.Text) & " ", con, adOpenDynamic, adLockPessimistic
If .RecordCount >= 1 Then
.Close
Dim strsql
strsql = "UPDATE  fees_stru set Tution_Fees = '" & feesstu.Textf.Text & _
        "', General_Fund= '" & Texgf.Text & "', Annual_Charges= '" & Texac.Text & _
        "', Examination_Fee= '" & Texef.Text & "' , Computer_Fee= '" & Texcf.Text & _
        "', Admission_Fee= '" & Texaf.Text & _
        "', Total= '" & TexfeeTot.Text & _
        "' where Class_ID =" & Val(Comfsid.Text) & ""
       
       con.Execute strsql
       MsgBox " data updated"
       
Else

    With rs_feestruct
        .AddNew
        .Fields("Std") = feesstu.Texstd.Text
        .Fields("Div") = feesstu.Texdiv.Text
        .Fields("Class_ID") = feesstu.Comfsid.Text
        .Fields("Tution_Fees") = feesstu.Textf.Text
        .Fields("General_Fund") = feesstu.Texgf.Text
        .Fields("Annual_Charges") = feesstu.Texac.Text
        .Fields("Examination_Fee") = feesstu.Texef.Text
        .Fields("Computer_Fee") = feesstu.Texcf.Text
        .Fields("Admission_Fee") = feesstu.Texaf.Text
        .Fields("Total") = feesstu.TexfeeTot.Text
        .Update
    End With
    MsgBox "Fee Structure Successfully Added", vbInformation + vbOKOnly, "Fee Structure"
    cmdfee_mod.Enabled = True
    cmdfee_sav.Enabled = False
End If
End With
End Sub

Private Sub cmdFeeSturOk_Click()
Unload Me
End Sub

Private Sub Comfsid_Click()
Textf.Text = ""
Texgf.Text = ""
Texac.Text = ""
Texef.Text = ""
Texcf.Text = ""
Texaf.Text = ""
TexfeeTot.Text = ""
Call connect
 
 With rs_feesstrufind
    If .State = adStateOpen Then .Close
    .Open "select Std,Div from class_mstr where Class_ID = " & Val(Comfsid.Text) & " ", con, adOpenDynamic, adLockPessimistic
    Do Until .EOF
    Texstd.Text = .Fields("Std").Value
    Texdiv.Text = .Fields("Div").Value
    .MoveNext
    Loop
    .Close
 End With


With rs_feesstrufind
 If .State = adStateOpen Then .Close
 .Open "select Tution_Fees,General_Fund,Annual_Charges,Examination_Fee,Computer_Fee,Admission_Fee,Total from fees_stru where fees_stru.Class_ID = " & Val(Comfsid.Text) & " ", con, adOpenDynamic, adLockPessimistic
If .RecordCount <= 0 Then
MsgBox " no fees structure for this class ADD new"
 End If
 Do Until .EOF
 'Texstd.Text = .Fields("Std").Value
 Textf.Text = .Fields("Tution_Fees").Value
 Texgf.Text = .Fields("General_Fund").Value
 Texac.Text = .Fields("Annual_Charges").Value
 Texef.Text = .Fields("Examination_Fee").Value
 Texcf.Text = .Fields("Computer_Fee").Value
 Texaf.Text = .Fields("Admission_Fee").Value
 TexfeeTot.Text = .Fields("Total").Value
 .MoveNext
 Loop
 .Close
 End With


End Sub


Private Sub Form_Load()
Call connect
With rs_feesstrufind
If .State = adStateOpen Then .Close
.Open "select Class_ID from class_mstr", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Comfsid.AddItem .Fields("Class_ID")
.MoveNext
Loop
.Close
End With
Call CenterForm(Me)
End Sub

Private Sub TexAc_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)
End Sub

Private Sub TexAc_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub TexAf_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)
End Sub

Private Sub TexAf_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub TexCf_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)

End Sub

Private Sub TexCf_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub TexEf_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)
End Sub

Private Sub TexEf_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
Private Sub TexGf_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)
End Sub

Private Sub Texgf_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Textf_Change()
TexfeeTot.Text = Val(Texcf.Text) + Val(Texef.Text) + Val(Texgf.Text) + Val(Texac.Text) + Val(Textf.Text) + Val(Texaf.Text)
End Sub

Private Sub TexTf_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
