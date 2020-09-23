VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStudFull 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student Detail"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   9000
      TabIndex        =   80
      Top             =   2160
      Width           =   1455
      Begin VB.CommandButton cmdxt 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Close"
         Height          =   495
         Left            =   240
         MouseIcon       =   "frmStudFull.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdCan 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         Height          =   495
         Left            =   240
         MouseIcon       =   "frmStudFull.frx":0152
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdSav 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         MouseIcon       =   "frmStudFull.frx":02A4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox txtvf 
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
      Left            =   2760
      TabIndex        =   72
      Top             =   720
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      MouseIcon       =   "frmStudFull.frx":03F6
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "frmStudFull.frx":0CD0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape7"
      Tab(0).Control(1)=   "Shape6"
      Tab(0).Control(2)=   "Label30"
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(12)=   "Label4"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(15)=   "pcbox"
      Tab(0).Control(16)=   "txtln"
      Tab(0).Control(17)=   "txtmon"
      Tab(0).Control(18)=   "txtgn"
      Tab(0).Control(19)=   "txtfan"
      Tab(0).Control(20)=   "txtpa"
      Tab(0).Control(21)=   "txtmn"
      Tab(0).Control(22)=   "txtfn"
      Tab(0).Control(23)=   "Texcity"
      Tab(0).Control(24)=   "Texstate"
      Tab(0).Control(25)=   "Texpin"
      Tab(0).Control(26)=   "Texphone"
      Tab(0).Control(27)=   "Texmobile"
      Tab(0).Control(28)=   "Texemail"
      Tab(0).Control(29)=   "cmdchange"
      Tab(0).Control(30)=   "cdb"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Identity/Class"
      TabPicture(1)   =   "frmStudFull.frx":0CEC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label24"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label9"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtgen"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtpob"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtnat"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtage"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtrol"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtdoa"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtdiv"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtstd"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "DTPdob"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Family Information"
      TabPicture(2)   =   "frmStudFull.frx":0D08
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label32"
      Tab(2).Control(1)=   "Label33"
      Tab(2).Control(2)=   "Label34"
      Tab(2).Control(3)=   "Label35"
      Tab(2).Control(4)=   "Label36"
      Tab(2).Control(5)=   "Label38"
      Tab(2).Control(6)=   "Label39"
      Tab(2).Control(7)=   "Label40"
      Tab(2).Control(8)=   "Label41"
      Tab(2).Control(9)=   "Label42"
      Tab(2).Control(10)=   "Label43"
      Tab(2).Control(11)=   "Shape2"
      Tab(2).Control(12)=   "Label37"
      Tab(2).Control(13)=   "Shape3"
      Tab(2).Control(14)=   "Label31"
      Tab(2).Control(15)=   "txtrn"
      Tab(2).Control(16)=   "txtmt"
      Tab(2).Control(17)=   "txtdep"
      Tab(2).Control(18)=   "txtpo"
      Tab(2).Control(19)=   "txtan"
      Tab(2).Control(20)=   "txtph"
      Tab(2).Control(21)=   "txtrel"
      Tab(2).Control(22)=   "txtocc"
      Tab(2).Control(23)=   "txtct"
      Tab(2).Control(24)=   "Texcas"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Previous School Info"
      TabPicture(3)   =   "frmStudFull.frx":0D24
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label29"
      Tab(3).Control(1)=   "Label28"
      Tab(3).Control(2)=   "Label22"
      Tab(3).Control(3)=   "Label19"
      Tab(3).Control(4)=   "Shape1"
      Tab(3).Control(5)=   "Label20"
      Tab(3).Control(6)=   "txtclsps"
      Tab(3).Control(7)=   "txtaps"
      Tab(3).Control(8)=   "txtps"
      Tab(3).Control(9)=   "txtreg"
      Tab(3).ControlCount=   10
      Begin MSComCtl2.DTPicker DTPdob 
         Height          =   375
         Left            =   2040
         TabIndex        =   86
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   58916865
         CurrentDate     =   40251
      End
      Begin MSComDlg.CommonDialog cdb 
         Left            =   -68160
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdchange 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Change"
         Height          =   375
         Left            =   -68400
         MouseIcon       =   "frmStudFull.frx":0D40
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Texcas 
         Height          =   375
         Left            =   -73200
         TabIndex        =   84
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox Texemail 
         Height          =   375
         Left            =   -72120
         TabIndex        =   79
         Top             =   6960
         Width           =   2535
      End
      Begin VB.TextBox Texmobile 
         Height          =   375
         Left            =   -72120
         MaxLength       =   14
         TabIndex        =   78
         Top             =   6480
         Width           =   2535
      End
      Begin VB.TextBox Texphone 
         Height          =   375
         Left            =   -72120
         MaxLength       =   14
         TabIndex        =   77
         Top             =   6000
         Width           =   2535
      End
      Begin VB.TextBox Texpin 
         Height          =   375
         Left            =   -72120
         MaxLength       =   7
         TabIndex        =   76
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox Texstate 
         Height          =   375
         Left            =   -72120
         TabIndex        =   75
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox Texcity 
         Height          =   375
         Left            =   -72120
         TabIndex        =   74
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox txtreg 
         Height          =   375
         Left            =   -71040
         TabIndex        =   71
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtstd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   70
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtdiv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   69
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtdoa 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   68
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtct 
         Height          =   375
         Left            =   -73200
         TabIndex        =   67
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtps 
         Height          =   405
         Left            =   -71760
         TabIndex        =   61
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtaps 
         Height          =   405
         Left            =   -71760
         TabIndex        =   60
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtclsps 
         Height          =   405
         Left            =   -71760
         TabIndex        =   59
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtocc 
         Height          =   405
         Left            =   -71400
         TabIndex        =   45
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtrel 
         Height          =   405
         Left            =   -71400
         TabIndex        =   44
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtph 
         Height          =   405
         Left            =   -73320
         TabIndex        =   43
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtan 
         Height          =   405
         Left            =   -71400
         TabIndex        =   42
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtpo 
         Height          =   405
         Left            =   -70080
         TabIndex        =   41
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtdep 
         Height          =   315
         Left            =   -71400
         TabIndex        =   40
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtmt 
         Height          =   405
         Left            =   -73200
         TabIndex        =   39
         Top             =   5640
         Width           =   2295
      End
      Begin VB.TextBox txtrn 
         Height          =   405
         Left            =   -73200
         TabIndex        =   38
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtrol 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtage 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtnat 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtpob 
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtgen 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtfn 
         Height          =   375
         Left            =   -72120
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtmn 
         Height          =   375
         Left            =   -72120
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtpa 
         Height          =   855
         Left            =   -72120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox txtfan 
         Height          =   375
         Left            =   -72120
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtgn 
         Height          =   375
         Left            =   -72120
         TabIndex        =   4
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtmon 
         Height          =   375
         Left            =   -72120
         TabIndex        =   3
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtln 
         Height          =   375
         Left            =   -72120
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Image pcbox 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   -68640
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Family Background "
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
         Left            =   -74400
         TabIndex        =   46
         Top             =   4200
         Width           =   2235
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   2895
         Left            =   -74760
         Top             =   4320
         Width           =   6615
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Parent/Gardian Information "
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
         Left            =   -74520
         TabIndex        =   52
         Top             =   600
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   3135
         Left            =   -74760
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Previous Information :"
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
         Left            =   -74400
         TabIndex        =   66
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   3135
         Left            =   -74640
         Top             =   1440
         Width           =   6255
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Address of previous School:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74400
         TabIndex        =   65
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Class/ STD in previous School"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74400
         TabIndex        =   64
         Top             =   3960
         Width           =   2445
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "School mention above was Recognise :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74400
         TabIndex        =   63
         Top             =   3240
         Width           =   3285
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of previous School:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74400
         TabIndex        =   62
         Top             =   2160
         Width           =   2100
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   58
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Home:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   57
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Rs"
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
         Left            =   -69480
         TabIndex        =   56
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation ship With student:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   55
         Top             =   1320
         Width           =   2265
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "annual Income:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   54
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Office:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -71280
         TabIndex        =   53
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Dependent in family"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   51
         Top             =   2760
         Width           =   2550
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   50
         Top             =   6120
         Width           =   840
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mother tongue :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   49
         Top             =   5640
         Width           =   1320
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cast :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   48
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74520
         TabIndex        =   47
         Top             =   4800
         Width           =   750
      End
      Begin VB.Label Label9 
         Caption         =   "Identity:"
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
         Left            =   360
         TabIndex        =   37
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll no:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   36
         Top             =   5520
         Width           =   675
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Admission:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   35
         Top             =   6000
         Width           =   1650
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DIV:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   34
         Top             =   5160
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Class:-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   33
         Top             =   4800
         Width           =   585
      End
      Begin VB.Label Label16 
         Caption         =   "Class/STD :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   4440
         Width           =   1320
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
         Left            =   960
         TabIndex        =   30
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
         Left            =   960
         TabIndex        =   29
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
         Left            =   960
         TabIndex        =   28
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
         Left            =   960
         TabIndex        =   27
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
         Left            =   960
         TabIndex        =   26
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label2 
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
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         TabIndex        =   20
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label4 
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
         TabIndex        =   19
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label5 
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
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         TabIndex        =   17
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         TabIndex        =   16
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label11 
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
         TabIndex        =   15
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label12 
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
         TabIndex        =   14
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label13 
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
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label14 
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
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         TabIndex        =   11
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label21 
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
         TabIndex        =   10
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label30 
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
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderWidth     =   2
         Height          =   3015
         Left            =   240
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   2055
         Left            =   240
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
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmStudFull.frx":0E92
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6240
      TabIndex        =   87
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label44 
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
      Left            =   960
      TabIndex        =   73
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Detailed Info :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   -240
      Picture         =   "frmStudFull.frx":0F1A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmStudFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pic_name As String, pic_ext As String, pic_changed As Boolean
Dim d As Date
Dim p As String

Private Sub cmdClassfindOk_Click()
Unload Me
End Sub

Private Sub cmdCan_Click()
Unload frmStudGeneral
Unload Me
End Sub

Private Sub cmdchange_Click()
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

Private Sub cmdSav_Click()
On Error GoTo handle
Call connect
d = DTPdob.Value
Dim strsql, strsql1, strsql2, strsql3, strsql4, strsql5, strsql6, strpic
strsql = "UPDATE  student_mstr set First_name='" & txtfn.Text & "',Middle_name='" & txtmn.Text & _
                 "',Last_name='" & txtln.Text & "',mother_name='" & txtmon.Text & _
                 "',father_name='" & txtfan.Text & "',name_of_gardian='" & txtgn.Text & _
                 "',Permanent_address='" & txtpa.Text & "',Std='" & txtstd.Text & _
                 "' where student_id =" & Val(txtvf.Text) & ""
strsql1 = "UPDATE  student_mstr set Div='" & txtdiv.Text & "',city='" & Texcity.Text & _
                 "',state='" & Texstate.Text & "',pincode='" & Texpin.Text & _
                 "',phone_no='" & Texphone.Text & "',mobile_no='" & Texmobile.Text & _
                 "',Email='" & Texemail.Text & "',gender='" & txtgen.Text & _
                 "',place_of_birth='" & txtpob.Text & "',DOB='" & d & _
                 "' where student_id =" & Val(txtvf.Text) & ""
strsql2 = "UPDATE  student_mstr set Nationality='" & txtnat.Text & "',Age='" & txtage.Text & _
                 "',Relationship_With_student='" & txtrel.Text & "',Occupation='" & txtocc.Text & _
                 "',annual_Income='" & txtan.Text & "',Number_of_Dependent_in_family='" & txtdep.Text & _
                 "',Phone_Home='" & txtph.Text & "',Phone_Office='" & txtpo.Text & _
                 "' where student_id =" & Val(txtvf.Text) & ""

strsql3 = "UPDATE student_mstr set Caste='" & Texcas.Text & "' where student_id=" & Val(txtvf.Text)

strsql4 = "UPDATE  student_mstr set Name_of_previous_School='" & txtps.Text & "',Address_of_previous_School='" & txtaps.Text & _
                 "',School_mention_above_was_Recognise='" & txtreg.Text & "',STD_in_previous_School='" & txtclsps.Text & _
                  "' where student_id =" & Val(txtvf.Text) & ""

strsql5 = "UPDATE  student_mstr set Mother_tongue='" & txtmt.Text & "',Category='" & txtct.Text & _
          "' where student_id =" & Val(txtvf.Text) & ""
strsql6 = "UPDATE  student_mstr set Religion='" & txtrn.Text & _
           "' where student_id =" & Val(txtvf.Text) & ""
           
If pic_name <> "" Then
FileCopy pic_name, App.Path & "\Miscellaneous\STUDENT_IMAGE\" & txtvf.Text & pic_ext
p = "\Miscellaneous\STUDENT_IMAGE\" & txtvf.Text & pic_ext
strpic = "UPDATE student_mstr set picture='" & p & "' where student_id=" & Val(txtvf.Text)
con.Execute strpic
End If

con.Execute strsql
con.Execute strsql1
con.Execute strsql2
con.Execute strsql4
con.Execute strsql3
con.Execute strsql5
con.Execute strsql6

MsgBox "data updated"
Exit Sub
handle:
MsgBox "Error :" & " " & Error$, vbCritical + vbOKOnly, "ERROR"
End Sub

Private Sub cmdxt_Click()
Unload frmStudGeneral
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
pic_name = ""
txtvf.Text = frmStudGeneral.txtsuid.Text
Call connect
With rs_find
Dim c As Double
c = (txtvf.Text)
If .State = adStateOpen Then .Close
.Open "select * from student_mstr where student_id = " & c & "", con, adOpenDynamic, adLockPessimistic
Do Until .EOF
txtfn.Text = .Fields("First_name").Value
txtmn.Text = .Fields("Middle_name").Value
txtln.Text = .Fields("Last_name").Value
txtmon.Text = .Fields("mother_name").Value
txtfan.Text = .Fields("father_name").Value
txtgn.Text = .Fields("name_of_gardian").Value
txtpa.Text = .Fields("Permanent_address").Value
txtstd.Text = .Fields("Std").Value
txtdiv.Text = .Fields("Div").Value
Texcity.Text = .Fields("city").Value
Texstate.Text = .Fields("state").Value
Texpin.Text = .Fields("pincode")
Texphone.Text = .Fields("phone_no")
Texmobile.Text = .Fields("mobile_no")
Texemail.Text = .Fields("Email")
txtgen.Text = .Fields("gender")
txtpob.Text = .Fields("place_of_birth")
DTPdob.Value = .Fields("DOB")
txtnat.Text = .Fields("Nationality")
txtage.Text = .Fields("Age")
txtrol.Text = .Fields("roll_no")
txtdoa.Text = .Fields("Date_of_Admission")
txtrel.Text = .Fields("Relationship_With_student")
txtocc.Text = .Fields("Occupation")
txtan.Text = .Fields("annual_Income")
txtdep.Text = .Fields("Number_of_Dependent_in_family")
txtph.Text = .Fields("Phone_Home")
txtpo.Text = .Fields("Phone_Office")
txtrn.Text = .Fields("Religion")
Texcas.Text = .Fields("Caste")
txtmt.Text = .Fields("Mother_tongue")
txtct.Text = .Fields("Category")
txtps.Text = .Fields("Name_of_previous_School")
txtaps.Text = .Fields("Address_of_previous_School")
txtreg.Text = .Fields("School_mention_above_was_Recognise")
txtclsps.Text = .Fields("STD_in_previous_School")
p = App.Path & .Fields("picture")
pcbox.Picture = LoadPicture(p)
.MoveNext
Loop
.Close
End With
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

Private Sub txtan_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub txtdep_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub txtph_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub
