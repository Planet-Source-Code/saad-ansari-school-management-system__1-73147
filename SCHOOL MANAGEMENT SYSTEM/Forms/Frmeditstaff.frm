VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frmeditstaff 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edit Staff Detail"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   10080
      TabIndex        =   102
      Top             =   1800
      Width           =   1335
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "Frmeditstaff.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdxt 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Frmeditstaff.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdCan 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Back"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Frmeditstaff.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdSav 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         MouseIcon       =   "Frmeditstaff.frx":03F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdt 
         BackColor       =   &H00C0E0FF&
         Caption         =   "E&dit"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Frmeditstaff.frx":0548
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox texsffid 
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   43
      Top             =   1440
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      MousePointer    =   99
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      MouseIcon       =   "Frmeditstaff.frx":069A
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "Frmeditstaff.frx":0F74
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cdb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_chgP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Texrel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Texage"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Texbp"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Texnat"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Texcivil"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Texcast"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Comsex"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Texemail"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Texmo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Texcn"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Texpin"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Texstate"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Texcity"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Texadd"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Texmname"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Texlname"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Texfname"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPdob"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label57"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label36"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label35"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label34"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label33"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label32"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label20"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label19"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label7"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label8"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label9"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label22"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label27"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label28"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label29"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label31"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "pcbox"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Acemedic Detail"
      TabPicture(1)   =   "Frmeditstaff.frx":0F90
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label37"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label38"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label39"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label40"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label41"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label42"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label43"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label45"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label46"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label47"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label48"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label49"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label50"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label51"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label52"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label53"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label54"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label55"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label56"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "DTPCT"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "DTPendate"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Comjob"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Comdepid"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Texsal"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Texenroll"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Texqual"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Texdepname"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Texsumqual"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Texpost"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Texrem"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Texpre"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Texnoyear"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Comexp"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Comtype"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Texsub4"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Texsub3"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Texsub2"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Texsub1"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      Begin MSComDlg.CommonDialog cdb 
         Left            =   -69600
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_chgP 
         Caption         =   "Change Picture"
         Height          =   375
         Left            =   -67800
         TabIndex        =   18
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Texrel 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   22
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox Texage 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   21
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox Texbp 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   20
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox Texnat 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   19
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox Texcivil 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   24
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox Texcast 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -68040
         TabIndex        =   23
         Top             =   5520
         Width           =   2655
      End
      Begin VB.TextBox Texsub1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   38
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox Texsub2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   39
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox Texsub3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   40
         Top             =   6360
         Width           =   1695
      End
      Begin VB.TextBox Texsub4 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   41
         Top             =   6840
         Width           =   1695
      End
      Begin VB.ComboBox Comtype 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frmeditstaff.frx":0FAC
         Left            =   6480
         List            =   "Frmeditstaff.frx":0FB6
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4440
         Width           =   2655
      End
      Begin VB.ComboBox Comexp 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frmeditstaff.frx":0FD2
         Left            =   7560
         List            =   "Frmeditstaff.frx":0FDC
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Texnoyear 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         TabIndex        =   34
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Texpre 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         TabIndex        =   36
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox Texrem 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   7560
         TabIndex        =   35
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Texpost 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         TabIndex        =   29
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox Texsumqual 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox Texdepname 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Texqual 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Texenroll 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         TabIndex        =   30
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Texsal 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2715
         TabIndex        =   32
         Top             =   5880
         Width           =   1335
      End
      Begin VB.ComboBox Comdepid 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2715
         TabIndex        =   25
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox Comjob 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frmeditstaff.frx":0FF6
         Left            =   2715
         List            =   "Frmeditstaff.frx":1000
         TabIndex        =   31
         Top             =   4920
         Width           =   1935
      End
      Begin VB.ComboBox Comsex 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frmeditstaff.frx":101A
         Left            =   -73170
         List            =   "Frmeditstaff.frx":1024
         TabIndex        =   10
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Texemail 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   17
         Top             =   7080
         Width           =   3015
      End
      Begin VB.TextBox Texmo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   16
         Top             =   6600
         Width           =   3015
      End
      Begin VB.TextBox Texcn 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   15
         Top             =   6120
         Width           =   3015
      End
      Begin VB.TextBox Texpin 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   14
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox Texstate 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   13
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox Texcity 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   12
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox Texadd 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -73170
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Texmname 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   8
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Texlname 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73170
         TabIndex        =   9
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Texfname 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73200
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPdob 
         Height          =   375
         Left            =   -73170
         TabIndex        =   44
         Top             =   2520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         Format          =   20709377
         CurrentDate     =   40251.5559837963
      End
      Begin MSComCtl2.DTPicker DTPendate 
         Height          =   375
         Left            =   2715
         TabIndex        =   75
         Top             =   5400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   40220
      End
      Begin MSComCtl2.DTPicker DTPCT 
         Height          =   375
         Left            =   2715
         TabIndex        =   76
         Top             =   6360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20709377
         CurrentDate     =   40220
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion :"
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
         Left            =   -69090
         TabIndex        =   101
         Top             =   5160
         Width           =   915
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status :"
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
         Left            =   -69360
         TabIndex        =   100
         Top             =   6000
         Width           =   1185
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age :"
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
         Left            =   -68760
         TabIndex        =   99
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality :"
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
         Left            =   -69315
         TabIndex        =   98
         Top             =   3720
         Width           =   1140
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place :"
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
         Left            =   -69345
         TabIndex        =   97
         Top             =   4080
         Width           =   1170
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caste :"
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
         Left            =   -69000
         TabIndex        =   96
         Top             =   5640
         Width           =   645
      End
      Begin VB.Label Label56 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5700
         TabIndex        =   95
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Label Label55 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5700
         TabIndex        =   94
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label54 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5700
         TabIndex        =   93
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label53 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5700
         TabIndex        =   92
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label52 
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
         Left            =   5760
         TabIndex        =   91
         Top             =   4440
         Width           =   585
      End
      Begin VB.Label Label51 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   90
         Top             =   1080
         Width           =   2130
      End
      Begin VB.Label Label50 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5640
         TabIndex        =   89
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label49 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5640
         TabIndex        =   88
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label48 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5640
         TabIndex        =   87
         Top             =   2760
         Width           =   1965
      End
      Begin VB.Label Label47 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2055
         TabIndex        =   86
         Top             =   4080
         Width           =   540
      End
      Begin VB.Label Label46 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   0
         TabIndex        =   85
         Top             =   3000
         Width           =   2595
      End
      Begin VB.Label Label45 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   900
         TabIndex        =   84
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label43 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   705
         TabIndex        =   83
         Top             =   1200
         Width           =   1890
      End
      Begin VB.Label Label42 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1035
         TabIndex        =   82
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label41 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1245
         TabIndex        =   81
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label40 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1470
         TabIndex        =   80
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Label Label39 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   405
         TabIndex        =   79
         Top             =   6480
         Width           =   2190
      End
      Begin VB.Label Label38 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   900
         TabIndex        =   78
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label37 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1845
         TabIndex        =   77
         Top             =   5880
         Width           =   750
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Age:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   74
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   73
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "place of birth"
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
         Left            =   -74040
         TabIndex        =   72
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   71
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   70
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   69
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "city "
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
         Left            =   -73920
         TabIndex        =   68
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "state "
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
         Left            =   -73920
         TabIndex        =   67
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   66
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent address "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73920
         TabIndex        =   65
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "pincode:"
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
         Left            =   -73920
         TabIndex        =   64
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "phone no: land line "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   63
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "mobile:"
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
         Left            =   -73920
         TabIndex        =   62
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "father name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   61
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "mother name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   60
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Guardian(If any) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73920
         TabIndex        =   59
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:-"
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
         Left            =   -73920
         TabIndex        =   58
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label44 
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
         Height          =   255
         Left            =   -73920
         TabIndex        =   57
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderWidth     =   2
         Height          =   3015
         Left            =   -74760
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   2055
         Left            =   -74760
         Top             =   4560
         Width           =   6135
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   7095
         Left            =   -74280
         Top             =   480
         Width           =   5295
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   7095
         Left            =   -74160
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
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
         Left            =   -73890
         TabIndex        =   56
         Top             =   7200
         Width           =   660
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile :"
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
         Left            =   -73995
         TabIndex        =   55
         Top             =   6720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pincode :"
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
         Left            =   -74115
         TabIndex        =   54
         Top             =   5760
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State :"
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
         Left            =   -73845
         TabIndex        =   53
         Top             =   5280
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City :"
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
         Left            =   -73710
         TabIndex        =   52
         Top             =   4800
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name :"
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
         Left            =   -74355
         TabIndex        =   51
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :"
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
         Left            =   -74370
         TabIndex        =   50
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name :"
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
         Left            =   -74610
         TabIndex        =   49
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex :"
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
         Left            =   -73725
         TabIndex        =   48
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of birth :"
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
         Left            =   -74520
         TabIndex        =   47
         Top             =   2640
         Width           =   1290
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number :"
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
         Left            =   -74880
         TabIndex        =   46
         Top             =   6240
         Width           =   1650
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
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
         Left            =   -74100
         TabIndex        =   45
         Top             =   3600
         Width           =   870
      End
      Begin VB.Image pcbox 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   -68040
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Label Label58 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Frmeditstaff.frx":1036
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6600
      TabIndex        =   103
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Staff ID:-"
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
      Height          =   375
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Staff Detail:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Frmeditstaff.frx":10BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Frmeditstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim d, d1, d2 As Date
Dim sub1, sub2, sub3, sub4 As String
Dim p As String
Public Sub Locked()
Texfname.Enabled = False
Texmname.Enabled = False
Texlname.Enabled = False
Texadd.Enabled = False
Texcity.Enabled = False
Texstate.Enabled = False
Texpin.Enabled = False
Texcn.Enabled = False
Texmo.Enabled = False
Texemail.Enabled = False
Texnat.Enabled = False
Texbp.Enabled = False
Texrel.Enabled = False
Texcast.Enabled = False
Texcivil.Enabled = False
Comtype.Enabled = False
Texdepname.Enabled = False
Texqual.Enabled = False
Texsumqual.Enabled = False
Texpost.Enabled = False
Comjob.Enabled = False
texsal.Enabled = False
DTPCT.Enabled = False
Comexp.Enabled = False
Texnoyear.Enabled = False
Texrem.Enabled = False
Texpre.Enabled = False

End Sub

Public Sub unlocked()

Texfname.Enabled = True
Texmname.Enabled = True
Texlname.Enabled = True
Texadd.Enabled = True
Texcity.Enabled = True
Texstate.Enabled = True
Texpin.Enabled = True
Texcn.Enabled = True
Texmo.Enabled = True
Texemail.Enabled = True
Texnat.Enabled = True
Texbp.Enabled = True
Texrel.Enabled = True
Texcast.Enabled = True
Texcivil.Enabled = True
Comtype.Enabled = True
Texdepname.Enabled = True
Texqual.Enabled = True
Texsumqual.Enabled = True
Texpost.Enabled = True
Comjob.Enabled = True
texsal.Enabled = True
DTPCT.Enabled = True
Comexp.Enabled = True
Texnoyear.Enabled = True
Texrem.Enabled = True
Texpre.Enabled = True

End Sub

Private Sub cmd_chgP_Click()
 
    cdb.Filter = "All Picture Files|*.jpg;*.gif;*.bmp;*.wmf;*.ico|JPEG Images(*.jpg)|*.jpg|Bitmap Images (*.bmp)|*.bmp|Word Meta Files (*.wmf)|*.wmf|GIF Images (*.gif)|*.gif"
    cdb.ShowOpen
    If cdb.FileName <> "" Then
        pcbox.Picture = LoadPicture(cdb.FileName)
        pic_name = cdb.FileName
        pic_ext = Right(cdb.FileTitle, 4)
        pic_changed = True
    End If
End Sub

Private Sub cmdCan_Click()
frmStffGeneral.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
Call Locked
cmdEdt.Enabled = True
cmdSav.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub cmdEdt_Click()
Call unlocked
cmdSav.Enabled = True
cmdCancel.Enabled = True
cmdEdt.Enabled = False
End Sub

Private Sub cmdSav_Click()


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
d = DTPdob.Value
d1 = DTPCT.Value
d2 = DTPendate.Value
Dim strsql, strsql1, strsql2, strpic As String
strsql = "UPDATE  staff_mstr set fname='" & Texfname.Text & _
        "',mname='" & Texmname.Text & "',lname='" & Texlname.Text & _
        "',dob='" & d & "',sex='" & Comsex.Text & _
        "',address='" & Texadd.Text & "',city='" & Texcity.Text & _
        "',state='" & Texstate.Text & "',pincode='" & Texpin.Text & _
        "',con_no='" & Texcn.Text & "',mob_no='" & Texmo.Text & _
        "',email='" & Texemail.Text & "',nationality='" & Texnat.Text & _
        "',birthplace='" & Texbp.Text & "',age='" & Texage.Text & _
        "' where staff_id =" & Val(texsffid.Text)
strsql1 = "UPDATE  staff_mstr set religion='" & Texrel.Text & _
        "',caste='" & Texcast.Text & "',civilstatus='" & Texcivil.Text & _
        "',type='" & Comtype.Text & "',department_id='" & comdepid.Text & _
        "',department_name='" & Texdepname.Text & "',qualification='" & Texqual.Text & _
        "',summaryofqual='" & Texsumqual.Text & "',post='" & Texpost.Text & _
        "',enroll_yr='" & Texenroll.Text & "',jobstatus='" & Comjob.Text & _
        "',enrollDate='" & d2 & "',salary='" & texsal.Text & _
        "',cont_term='" & d1 & "',pre_exp='" & Comexp.Text & _
        "' where staff_id =" & Val(texsffid.Text)
strsql2 = "UPDATE  staff_mstr set no_of_yr='" & Texnoyear.Text & _
        "',remark='" & Texrem.Text & "',pre_employr='" & Texpre.Text & _
        "',Subject1='" & Texsub1.Text & "',Subject2='" & Texsub2.Text & _
        "',Subject3='" & Texsub3.Text & "',Subject4='" & Texsub4.Text & _
        "' where staff_id =" & Val(texsffid.Text)
If pic_name <> "" Then
FileCopy pic_name, App.Path & "\Miscellaneous\STAFF_IMAGE\" & texsffid.Text & pic_ext
p = "\Miscellaneous\STAFF_IMAGE\" & texsffid.Text & pic_ext
strpic = "UPDATE staff_mstr set picture='" & p & "' where staff_id=" & Val(texsffid.Text)
con.Execute strpic
End If
con.Execute strsql
con.Execute strsql1
con.Execute strsql2


MsgBox "Record Updated"
Call Locked
cmdEdt.Enabled = True
cmdSav.Enabled = False
cmdCancel.Enabled = False
Exit Sub
handle:
MsgBox "Error : " & Error$
End Sub

Private Sub cmdxt_Click()
Unload frmStffGeneral
Unload Me
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

Private Sub Comtype_Click()

If Comtype.Text = "NON-TEACHING" Then
Texsub1.Text = "NiL"
Texsub2.Text = "NiL"
Texsub3.Text = "NiL"
Texsub4.Text = "NiL"
Texsub1.Enabled = False
Texsub2.Enabled = False
Texsub3.Enabled = False
Texsub4.Enabled = False
End If
If Comtype.Text = "TEACHING" Then
Texsub1.Text = sub1
Texsub2.Text = sub2
Texsub3.Text = sub3
Texsub4.Text = sub4
Texsub1.Enabled = True
Texsub2.Enabled = True
Texsub3.Enabled = True
Texsub4.Enabled = True
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
pic_name = ""
texsffid.Text = frmStffGeneral.comstffid.Text
With rs_find
If .State = adStateOpen Then .Close
.Open "select * from staff_mstr where staff_id = " & Val(texsffid.Text), con, adOpenDynamic, adLockPessimistic
Do Until .EOF
Texfname.Text = .Fields("fname")
Texmname.Text = .Fields("mname")
Texlname.Text = .Fields("lname")
DTPdob.Value = .Fields("dob")
Comsex.Text = .Fields("sex")
Texadd.Text = .Fields("address")
Texcity.Text = .Fields("city")
Texstate.Text = .Fields("state")
Texpin.Text = .Fields("pincode")
Texcn.Text = .Fields("con_no")
Texmo.Text = .Fields("mob_no")
Texemail.Text = .Fields("email")
Texnat.Text = .Fields("nationality")
Texbp.Text = .Fields("birthplace")
Texage.Text = .Fields("age")
Texrel.Text = .Fields("religion")
Texcast.Text = .Fields("caste")
Texcivil.Text = .Fields("civilstatus")
Comtype.Text = .Fields("type")
comdepid.Text = .Fields("department_id")
Texdepname.Text = .Fields("department_name")
Texqual.Text = .Fields("qualification")
Texsumqual.Text = .Fields("summaryofqual")
Texpost.Text = .Fields("post")
Texenroll.Text = .Fields("enroll_yr")
Comjob.Text = .Fields("jobstatus")
DTPendate.Value = .Fields("enrollDate")
texsal.Text = .Fields("salary")
DTPCT.Value = .Fields("cont_term")
Comexp.Text = .Fields("pre_exp")
Texnoyear.Text = .Fields("no_of_yr")
Texrem.Text = .Fields("remark")
Texpre.Text = .Fields("pre_employr")
Texsub1.Text = .Fields("Subject1")
Texsub2.Text = .Fields("Subject2")
Texsub3.Text = .Fields("Subject3")
Texsub4.Text = .Fields("Subject4")
'Dim a As String
p = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(p)
.MoveNext
Loop
.Close
End With
sub1 = Texsub1.Text
sub2 = Texsub2.Text
sub3 = Texsub3.Text
sub4 = Texsub4.Text
End Sub

Private Sub Texcn_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texmo_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texnoyear_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texpin_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Texsal_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

