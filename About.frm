VERSION 4.00
Begin VB.Form frmAboutILIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Library Information System 1.0"
   ClientHeight    =   2025
   ClientLeft      =   2820
   ClientTop       =   1680
   ClientWidth     =   5025
   Height          =   2430
   Left            =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Top             =   1335
   Width           =   5145
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgSmallLogo 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   120
      Picture         =   "About.frx":0000
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":1C02
      Height          =   1575
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line lneILIS 
      X1              =   1200
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmAboutILIS"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
