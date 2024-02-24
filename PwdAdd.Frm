VERSION 4.00
Begin VB.Form frmAddEditUsr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   1440
   ClientWidth     =   2895
   Height          =   2565
   Icon            =   "PwdAdd.frx":0000
   Left            =   0
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Top             =   1095
   Width           =   3015
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   2
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Created"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "User name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmAddEditUsr"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmAddEditUsr
'***************************************************************************


Private Sub cmdAddEdit_Click(Index As Integer)
'EncryptPassword
  Dim Now As String
  On Error Resume Next
  Select Case Index
    Case 0
    If txtText(1).Text = "" Or txtText(2).Text = "" Then
        MsgBox "You can't leave user name and password blank", vbExclamation, strSystem
        Exit Sub
    ElseIf txtText(1).Text = "Admin" Or txtText(1).Text = "admin" Then
        MsgBox "You can't replace default admin", vbExclamation, strSystem
        Exit Sub
    End If
    
    
    SaveSetting strExecName, "Name", txtText(1).Text, txtText(0).Text
    SaveSetting strExecName, "Created", txtText(1).Text, txtText(3).Text
    SaveSetting strExecName, "User name", txtText(1).Text, txtText(2).Text ' encryption password removed
    SaveSetting strExecName, "Last log", txtText(1).Text, Now
    
    ' Should use .Add, this is SLOWWW..
    frmAdminPwd.updateGetAllSettings
    Unload Me
    Case 1
      Unload Me
  End Select
End Sub

Private Sub Form_Load()
      txtText(3).Text = Now
      Caption = "Add User"
End Sub
