VERSION 5.00
Begin VB.Form frmSidePanel 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   13995
   ClientLeft      =   195
   ClientTop       =   0
   ClientWidth     =   3300
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13995
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   120
      Picture         =   "frmSidePanel.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Staff Registration"
      Top             =   3480
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3120
      Left            =   0
      Picture         =   "frmSidePanel.frx":8C81
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3315
   End
End
Attribute VB_Name = "frmSidePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image2_Click()
frmsturegister1.Show
End Sub
