VERSION 4.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4335
   ClientLeft      =   3555
   ClientTop       =   2430
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4740
   Left            =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6375
   Top             =   2085
   Width           =   6495
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is protected by international copyright laws as described in Help About."
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY INFORMATION SYSTEM"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'===========================================================================
' frmAbout
' Author        :
' Last Updated  :
' Notes         :
'===========================================================================

Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii%)
    Unload Me
End Sub

Private Sub Form_Load()
    ' Make current form Centered
    CenterForm Me
End Sub

Private Sub lblInfo_Click(Index As Integer)
    Unload Me
End Sub

