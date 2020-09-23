VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm SchoolMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Life English School"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "SchoolMain.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "SchoolMain.frx":2C5A
   Moveable        =   0   'False
   Picture         =   "SchoolMain.frx":2F64
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3600
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      Height          =   10335
      Left            =   0
      Picture         =   "SchoolMain.frx":13266
      ScaleHeight     =   10275
      ScaleWidth      =   3135
      TabIndex        =   4
      Top             =   0
      Width           =   3195
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "    Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   8400
         Width           =   2775
         Begin VB.Image Image2 
            Height          =   240
            Left            =   120
            Picture         =   "SchoolMain.frx":229A8
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   11
            Top             =   840
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "      Menu Explorer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   5160
         Width           =   2775
         Begin VB.CommandButton cmd_xt 
            BackColor       =   &H00C0E0FF&
            Caption         =   "E&xit"
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
            Left            =   240
            MouseIcon       =   "SchoolMain.frx":22D32
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdChgPass 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Change Password"
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
            Left            =   240
            MouseIcon       =   "SchoolMain.frx":22E84
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdLogf 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Log Off"
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
            Left            =   1320
            MouseIcon       =   "SchoolMain.frx":22FD6
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   840
            Width           =   1095
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   120
            Picture         =   "SchoolMain.frx":23128
            Top             =   0
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lvMenu 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7858
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2400
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":234B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":24E44
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":267D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":28168
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":28A44
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":29720
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":2A004
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "SchoolMain.frx":2A8E0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "     Logged On as"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   6600
         Width           =   2775
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   75
         End
         Begin VB.Label lblp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preveilege :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1035
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   120
            Picture         =   "SchoolMain.frx":2C274
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   75
         End
      End
      Begin VB.Label NLESLBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME TO NLES AUTOMATED SYSTEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2640
         TabIndex        =   17
         Top             =   120
         Width           =   5400
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   4680
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10335
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Status : "
            TextSave        =   "Status : "
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "12:57 PM Saad"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "5/18/2010"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdminn 
         Caption         =   "Administration"
         Begin VB.Menu mnuSysUser 
            Caption         =   "System Users"
         End
         Begin VB.Menu mnuflnewUser 
            Caption         =   "New User"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnurepuser 
            Caption         =   "User Report"
         End
      End
      Begin VB.Menu mnutsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "&SystemLock"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnulogout 
         Caption         =   "&Logout"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuStudent 
         Caption         =   "Student"
         Begin VB.Menu mnustdtreg 
            Caption         =   "Student Enroll"
         End
         Begin VB.Menu mnuresultinfo 
            Caption         =   "Create Result"
         End
         Begin VB.Menu mnuStudel 
            Caption         =   "Delete Student"
         End
         Begin VB.Menu mnu_stu_att 
            Caption         =   "Student Attendance"
         End
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
         Begin VB.Menu mnustffreg 
            Caption         =   "Staff Enroll"
         End
         Begin VB.Menu mnuDlyAttn 
            Caption         =   "Daily Attendance"
         End
         Begin VB.Menu mnuSaly 
            Caption         =   "Salary"
         End
         Begin VB.Menu mnuStffsusp 
            Caption         =   "Delete Staff"
         End
      End
      Begin VB.Menu mnuEdsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepart 
         Caption         =   "Department"
         Begin VB.Menu mnuNewDepart 
            Caption         =   "New Department"
         End
         Begin VB.Menu mnuedtdep 
            Caption         =   "Edit Department"
         End
      End
      Begin VB.Menu mnuFee 
         Caption         =   "Fee"
         Begin VB.Menu mnuFees 
            Caption         =   "Fees Payment"
         End
         Begin VB.Menu mnuFeeStruct 
            Caption         =   "Fee Structure"
         End
      End
      Begin VB.Menu mnuClass 
         Caption         =   "Class"
         Begin VB.Menu mnuClsInfo 
            Caption         =   "Add Class"
         End
         Begin VB.Menu mnu_edit_class 
            Caption         =   "Edit class"
         End
         Begin VB.Menu mnuFindclass 
            Caption         =   "Find Class"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStudInfo 
         Caption         =   "Student Information"
         Begin VB.Menu mnuStudGenInfo 
            Caption         =   "General Info"
         End
         Begin VB.Menu mnuStudFee 
            Caption         =   "Fee Record"
         End
         Begin VB.Menu mnustuatt 
            Caption         =   "Attendance"
         End
      End
      Begin VB.Menu mnuVuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnustfInfo 
         Caption         =   "Staff Information"
         Begin VB.Menu mnuGenInfo 
            Caption         =   "General Info"
         End
         Begin VB.Menu mnustffatt 
            Caption         =   "Attendance"
         End
      End
      Begin VB.Menu mnuViewsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDepart 
         Caption         =   "Department"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuRepStud 
         Caption         =   "Student"
         Begin VB.Menu mnurepstu 
            Caption         =   "Information"
         End
         Begin VB.Menu mnuFeesrep 
            Caption         =   "Fees"
            Begin VB.Menu mnurepfee 
               Caption         =   "Fees Information (All)"
            End
            Begin VB.Menu mnuIndFee 
               Caption         =   "Fees Information (Individual)"
            End
         End
         Begin VB.Menu mnurepRes 
            Caption         =   "Result"
         End
      End
      Begin VB.Menu mnuRepStff 
         Caption         =   "Staff"
         Begin VB.Menu mnustffInfo 
            Caption         =   "Staff Info"
         End
         Begin VB.Menu mnuSalInfo 
            Caption         =   "Salary Info"
         End
      End
      Begin VB.Menu mnurepClass 
         Caption         =   "Class"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utilties"
      Begin VB.Menu mnuNtpad 
         Caption         =   "Notepad"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBackDB 
         Caption         =   "Backup DB"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhlpAbt 
         Caption         =   "About"
      End
      Begin VB.Menu mnuAck 
         Caption         =   "Acknowledgement"
      End
      Begin VB.Menu mnuhlpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortcuts 
         Caption         =   "Help on shortcuts"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "HelpOnline!"
      End
   End
End
Attribute VB_Name = "SchoolMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub open_Click()
frmsturegister1.Show

End Sub

Private Sub Command1_Click()
'frmPdf.Show
End Sub

Private Sub cmd_xt_Click()
Call mnuquit_Click
End Sub

Private Sub cmdCp_Click()
Shell "c:\windows\system32\control.exe", vbNormalFocus
End Sub

Private Sub cmdChgPass_Click()
frmChgPass.Show
End Sub

Private Sub cmdLogf_Click()
Call mnulogout_Click
End Sub

Private Sub MDIForm_Initialize()
Label5.Caption = Date
horizontal.Show
Vertical.Show
'=====Code for the icon of the listview
'========
With lvMenu
       Set .SmallIcons = ImageList1
       Set .Icons = ImageList1
        'For Sales
        .ListItems.Add , "user", "View Users", 1, 1
        .ListItems.Add , "Add Student", "Add New Student", 2, 2
        .ListItems.Add , "Edit Student Information", "Edit Student Information", 3, 3
        .ListItems.Add , "Result", "Result", 4, 4
        .ListItems.Add , "View Student Report", "View Student Report", 5, 5
        .ListItems.Add , "Result Report", "Result Report", 6, 6
        .ListItems.Add , "Fees Payment", "Fees Payment", 7, 7
        .ListItems.Add , "School Info", "School Info", 8, 8

    End With

End Sub
Private Sub lvMenu_DblClick()
Select Case lvMenu.SelectedItem.Key

Case "user": If SchoolMain.Label8.Caption = "Administrator" Then frmSysUser.Show
Case "Add Student": frmsturegister1.Show
Case "Edit Student Information": frmStudGeneral.Show
Case "Result": frmResult.Show
Case "Fees Payment": Feespayment.Show
Case "View Student Report": Call mnurepstu_Click
Case "School Info": frmSchoolInfo.Show
Case "Result Report": frmResultRep.Show
End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo handle
       If SchoolMain.Label2.Caption <> "" Then
         MsgBox " I N I T I A T E D   A P P L I C A T I O N    S H U T D O W N"
         
            Call connect
         
            With rs_find
            If .State = adStateOpen Then .Close
            .Open "select * from UserLog", con, adOpenDynamic, adLockPessimistic
            .AddNew
            .Fields("UserID").Value = SchoolMain.Label2.Caption
            .Fields("SessionStart").Value = frmLogin.lblval.Caption
            .Fields("SessionEnd").Value = Now
            .Fields("Description") = "Session Successfully Terminated"
            .Update
           .Close
        
           End With
            MsgBox "THANK YOU FOR USING NLES AUTOMATED SYSTEM" & vbCrLf & vbCrLf & "  SESSION ENDS AT : " & Now, vbInformation + vbOKOnly, "SESSION SUCCESSFULLY TERMINATED"
            End
        Else
            End
        End If
    
    


Exit Sub
handle:
MsgBox "Error : " & Error$, vbCritical + vbOKOnly, "ERROR"
End Sub

Private Sub mnu_edit_class_Click()
editclass.Show
End Sub

Private Sub mnu_stu_att_Click()
frm_stu_attendance.Show
End Sub

Private Sub mnuAck_Click()
frmAck.Show
End Sub

Private Sub mnuBackDB_Click()
BackupDatabase.Show
End Sub

Private Sub mnuCalc_Click()
On Error GoTo Err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "Calculator Missing"
End Sub

Private Sub mnuClsInfo_Click()
Addclass.Show
End Sub

Private Sub mnucontact_Click()
MsgBox "Please for any Query " & vbCrLf & "----------------------------- " & vbCrLf & "Just e-mail me at : saadansari12345@gmail.com" & vbCrLf & vbCrLf & " Just e-mail me at : sms.shakeal@gmail.com", vbOKOnly, "Online Help!"
End Sub

Private Sub mnuDlyAttn_Click()
frmstffattendance.Show
End Sub

Public Sub mnuedtdep_Click()
frmedtDep.Show
End Sub

Private Sub mnuFeeS_Click()
Feespayment.Show
End Sub

Private Sub mnuFeeStruct_Click()
feesstu.Show
End Sub

Private Sub mnuFindclass_Click()
Findclass.Show
End Sub

Private Sub mnuflnewUser_Click()
frmUserAdd.Show
End Sub

Private Sub mnuGenInfo_Click()
frmStffGeneral.Show
End Sub

Private Sub mnuhlpAbt_Click()
frmAbout.Show
End Sub

Private Sub mnuIndFee_Click()
frmIndfees.Show
End Sub

Private Sub mnuLock_Click()
Load frmLock
frmLock.Show vbModal
End Sub

Public Sub mnulogout_Click()
Dim rep
If SchoolMain.mnulogout.Caption = "Logout" Then
    rep = MsgBox("Are you sure you want to Logout ?", vbYesNo, "Confirm Logout")
        If rep = vbYes Then
            Call connect
            With rs_find
            If .State = adStateOpen Then .Close
            .Open "select * from UserLog", con, adOpenDynamic, adLockPessimistic
            .AddNew
            .Fields("UserID").Value = SchoolMain.Label2.Caption
            .Fields("SessionStart").Value = frmLogin.lblval.Caption
            .Fields("SessionEnd").Value = Now
            .Fields("Description") = "Session Successfully Terminated"
            .Update
            .Close
            End With
                      
            
            SchoolMain.Label2.Caption = ""
            SchoolMain.Label5.Caption = ""
            SchoolMain.Label6.Caption = ""
            SchoolMain.Label8.Caption = ""
            SchoolMain.Timer1.Enabled = False
            SchoolMain.StatusBar1.Panels(1).Text = SchoolMain.StatusBar1.Panels(1).Text + " "
            MsgBox "THANK YOU FOR USING NLES AUTOMATED SYSTEM" & vbCrLf & vbCrLf & "  SESSION ENDS AT : " & Now, vbInformation + vbOKOnly, "SUCCESSFULLY LOGGED OUT"
            
            
            Unload frmsturegister1
            Unload frmsturegister2
            Unload Addclass
            Unload BackupDatabase
            Unload editclass
            Unload Feespayment
            Unload feesstu
            Unload Findclass
            Unload frm_stu_attendance
            Unload frmAbout
            Unload frmAck
            Unload frmChgPass
            Unload frmclassrep
            Unload frmDepartment
            Unload Frmeditstaff
            Unload frmedtDep
            Unload frmHlp
            Unload frmIndfees
            Unload frmNewDepart
            Unload frmResult
            Unload frmResultRep
            Unload frmSchoolInfo
            Unload frmstffattendance
            Unload frmStffGeneral
            Unload frmStffSalary
            Unload frmstudel
            Unload frmStudFull
            Unload frmStudGeneral
            Unload frmSysUser
            Unload frmUserAdd
            Unload horizontal
            Unload Stu_att_detail
            Unload teach_att_detail
            Unload teachregister
            Unload teachregister2
            Unload teachsuspend
            Unload Vertical
            Unload classreport
            Unload druser
            Unload IndFees
            Unload ResultReport
            Unload StaffReport
            Unload StudentFees
            Unload StudentReport
                        
            SchoolMain.Enabled = False
            frmLogin.Show
            frmLogin.txtuid.Text = ""
            frmLogin.txtpassword.Text = ""
            
        Else
            GoTo dude
    End If

ElseIf SchoolMain.mnulogout.Caption = "Login" Then
    Call Login
End If
dude:
Exit Sub
End Sub

Public Sub mnuNewDepart_Click()
frmNewDepart.Show
End Sub

Private Sub mnuNtpad_Click()
On Error GoTo Err
    Shell "Notepad.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have Notepad installed in your computer.", vbExclamation, "Notepad Missing"
End Sub

Public Sub mnuquit_Click()
Dim rep
On Error GoTo handle
rep = MsgBox("Are you sure you wanna Quit ? ", vbExclamation + vbYesNo, "Program Exit")
If rep = vbYes Then
         If SchoolMain.Label2.Caption <> "" Then
         MsgBox " I N I T I A T E D   A P P L I C A T I O N    S H U T D O W N"
         
         Call connect
         
            With rs_find
            If .State = adStateOpen Then .Close
            .Open "select * from UserLog", con, adOpenDynamic, adLockPessimistic
            .AddNew
            .Fields("UserID").Value = SchoolMain.Label2.Caption
            .Fields("SessionStart").Value = frmLogin.lblval.Caption
            .Fields("SessionEnd").Value = Now
            .Fields("Description") = "Session Successfully Terminated"
            .Update
           .Close
           End With
        End If
        MsgBox "THANK YOU FOR USING NLES AUTOMATED SYSTEM" & vbCrLf & vbCrLf & "  SESSION ENDS AT : " & Now, vbInformation + vbOKOnly, "SESSION SUCCESSFULLY TERMINATED"
        End
    Else
        Exit Sub
End If

Exit Sub
handle:
MsgBox "Error : " & Error$, vbCritical + vbOKOnly, "ERROR"
End Sub


Private Sub mnurepClass_Click()
frmclassrep.Show
End Sub

Private Sub mnurepfee_Click()
AdodbConnection.Init_studFees_Report
End Sub

Private Sub mnurepRes_Click()
frmResultRep.Show
End Sub


Private Sub mnurepstu_Click()
AdodbConnection.Init_student_Report
End Sub

Private Sub mnurepuser_Click()
AdodbConnection.Init_user_Report
End Sub

Private Sub mnuresinfo_Click()
frmResult.Show
End Sub

Private Sub mnuresultinfo_Click()
frmResult.Show
End Sub

Private Sub mnuSalInfo_Click()
frm_staff_sal_rep.Show
End Sub

Private Sub mnuSaly_Click()
frmStffSalary.Show
End Sub

Private Sub mnuShortcuts_Click()
frmHlp.Show
End Sub

Private Sub mnustdtreg_Click()
frmsturegister1.Show
End Sub

Private Sub mnustffatt_Click()
teach_att_detail.Show
End Sub

Private Sub mnustffInfo_Click()
AdodbConnection.Init_staff_Report
End Sub

Private Sub mnustffreg_Click()
teachregister.Show
End Sub

Private Sub mnuStffsusp_Click()
teachsuspend.Show
End Sub

Private Sub mnustuatt_Click()
Stu_att_detail.Show
End Sub

Private Sub mnuStudel_Click()
frmstudel.Show
End Sub

Private Sub mnuStudFee_Click()
AdodbConnection.Init_studFees_Report
End Sub

Private Sub mnuStudGenInfo_Click()
frmStudGeneral.Show
End Sub

Private Sub mnuSysUser_Click()
frmSysUser.Show
End Sub

Private Sub mnuViewDepart_Click()
frmDepartment.Show
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Time$
End Sub

Private Sub Timer2_Timer()
NLESLBL.Left = NLESLBL.Left - 100 'code to move the newlife tag in the form

If NLESLBL.Left + NLESLBL.Width <= 0 Then
NLESLBL.Left = Picture1.Width
End If

End Sub

