VERSION 4.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1920
   ClientLeft      =   1995
   ClientTop       =   2265
   ClientWidth     =   5280
   Height          =   2325
   Icon            =   "PwdOptn.frx":0000
   Left            =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Top             =   1920
   Width           =   5400
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "&OK"
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame 
      Caption         =   "Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox chkOptions 
         Caption         =   "Require alphanumeric System password"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Minimum System password length"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtLength 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.VScrollBar vscPwdLen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         Max             =   1
         Min             =   8
         TabIndex        =   3
         Top             =   720
         Value           =   1
         Width           =   255
      End
      Begin VB.Label lbl 
         Caption         =   "Length"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmOptions
'***************************************************************************


Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
    Case 0  ' Set AlphaNumeric
    If chkOptions(0).Value = 1 Then
      boolSetPwdAlphaNum = True
    Else
      boolSetPwdAlphaNum = False
    End If
    Case 1  ' Set MinPwdLen
      If chkOptions(1).Value = 1 Then
        txtLength.Enabled = True
        vscPwdLen.Enabled = True
        lbl.Enabled = True
        boolSetPwdLen = True
        intSetPwdLen = CInt(txtLength.Text)
      Else
        txtLength.Enabled = False
        vscPwdLen.Enabled = False
        lbl.Enabled = False
        boolSetPwdLen = False
      End If
  End Select
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)
  Select Case Index
    Case 0
      SaveSetting appname:=strExecName, section:="Settings", Key:="AlphaNumeric", setting:=CInt(boolSetPwdAlphaNum)
      SaveSetting appname:=strExecName, section:="Settings", Key:="RequireMinPwd", setting:=CInt(boolSetPwdLen)
      SaveSetting appname:=strExecName, section:="Settings", Key:="MinPwdVal", setting:=CStr(txtLength.Text)
      Unload Me
    Case 1
      Unload Me
  End Select
End Sub
Private Sub Form_Load()
  ' Get System settings for options from registry; absolute them for boolean value
  chkOptions(0).Value = Abs(GetSetting(appname:=strExecName, section:="Settings", Key:="AlphaNumeric"))
  chkOptions(1).Value = Abs(GetSetting(appname:=strExecName, section:="Settings", Key:="RequireMinPwd"))
  txtLength.Text = GetSetting(appname:=strExecName, section:="Settings", Key:="MinPwdVal")
End Sub
' Change value according Vertical Scroll Bar
Private Sub vscPwdLen_Change()
  txtLength.Text = vscPwdLen.Value
End Sub
