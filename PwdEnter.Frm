VERSION 4.00
Begin VB.Form frmPwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter System Password"
   ClientHeight    =   1665
   ClientLeft      =   1680
   ClientTop       =   2340
   ClientWidth     =   6840
   Height          =   2070
   Icon            =   "PwdEnter.frx":0000
   Left            =   1620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Top             =   1995
   Width           =   6960
   Begin VB.TextBox txtUsr 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type a user name and password to identify yourself to the System."
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblLabel 
      Caption         =   "User name:"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   120
      Picture         =   "PwdEnter.frx":000C
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblLabel 
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Caption         =   "Password:"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmPwd"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


'***************************************************************************
' frmPwd
'***************************************************************************

Option Explicit
Public intTrialLogin        As Integer
Public strTrialLogin        As String
Public loginState           As Boolean


Private Sub cmdOKCancel_Click(Index As Integer)
    If gfLogConfirm = False Then currUsr = txtUsr.Text
    Select Case Index
        Case 0
            passwdLog
        Case 1
        If gfLogConfirm = False Then
            If MsgBox("Quit Application", vbYesNo + vbQuestion, strSystem) = vbYes Then
                End
            End If
        ElseIf gfLogConfirm = True Then
            Unload Me
        End If
    End Select
End Sub


Private Sub Form_Load()
  intTrialLogin = 1
  loginState = False
  strPasswd$ = GetSetting(appname:=strExecName, section:="User name", _
    Key:="Admin", Default:="admin")
  If strPasswd$ = "admin" Then
    txtUsr.Text = "Admin"
    txtUsr.Enabled = False
    lblLabel(4).Caption = lblLabel(4).Caption + ". Since this is first Log to System, just hit OK and change the password later."
    regPwdExer8                   ' register Pwd in registry
  End If
End Sub


Sub passwdLog()
'DecryptPassword
  Dim PwdSettings As Variant
  Dim Settings As Integer
    PwdSettings = GetAllSettings(appname:=strExecName, section:="User name")
    For Settings = LBound(PwdSettings, 1) To UBound(PwdSettings, 1)
      strUsrName$ = PwdSettings(Settings, 0)
      strPasswd$ = PwdSettings(Settings, 1) ' password decryption removed
      Debug.Print "PwdSettings(Settings, 1) - ", PwdSettings(Settings, 1)
      Debug.Print "getCurrUsrSecDate - ", getCurrUsrSecDate
      Debug.Print "strPasswd$ - ", strPasswd$
      If (txtUsr.Text = strUsrName$) And (txtPwd.Text = strPasswd$) Then loginState = True
    Next Settings
  Select Case loginState
  Case True
    currUsr = txtUsr.Text
    loginSuccess
  Case False
    loginFail
    intTrialLogin = intTrialLogin + 1
  End Select
End Sub


Sub loginSuccess()
    If currUsr <> "Admin" Then
        disableAdminControl
    Else
        enableAdminControl
    End If
    
    If gfLogConfirm = False Then
        frmMDIMainMenu.Caption = strSystem & " Logged as " & currUsr & " Win95 Logged : " & UserID
        ShowFirstGrp
        Unload Me
    Else
        Unload Me
        frmAddEditUsr.Show vbModal
    End If

End Sub


Public Sub disableAdminControl()
    frmMDIMainMenu.mnuUtility(1).Enabled = 0
End Sub


Public Sub enableAdminControl()
    frmMDIMainMenu.mnuUtility(1).Enabled = 1
End Sub


Sub loginFail()
    If intTrialLogin > 2 Then
        MsgBox "Invalid Access", vbCritical, strSystem
        If gfLogConfirm = False Then
            End
        Else
            Unload Me
        End If
    Else
        If gfLogConfirm = False Then
            MsgBox "Login fail " & intTrialLogin & " of 3. Please try again.", , strSystem
        Else
            Unload Me
        End If
    End If
End Sub


' Jump to passwdLog directly after hit enter
Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        passwdLog
    End If
End Sub


' Register PwdExer8b for first time
Public Sub regPwdExer8()
  SaveSetting appname:=strExecName, section:="User name", Key:="Admin", setting:=""
  SaveSetting appname:=strExecName, section:="Name", Key:="Admin", setting:=""
  SaveSetting appname:=strExecName, section:="Created", Key:="Admin", setting:=Now
  SaveSetting appname:=strExecName, section:="Last Log", Key:="Admin", setting:=Now
  
  ' Set default settings for System
  SaveSetting appname:=strExecName, section:="Settings", Key:="AlphaNumeric", setting:="0"
  SaveSetting appname:=strExecName, section:="Settings", Key:="RequireMinPwd", setting:="0"
  SaveSetting appname:=strExecName, section:="Settings", Key:="MinPwdVal", setting:="1"
End Sub


Private Sub txtUsr_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtPwd.SetFocus
    KeyAscii = 0
  End If
End Sub

