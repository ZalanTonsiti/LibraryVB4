VERSION 4.00
Begin VB.Form frmConfirmChangePwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change SystemName Password"
   ClientHeight    =   1560
   ClientLeft      =   4170
   ClientTop       =   1455
   ClientWidth     =   5370
   Height          =   1965
   Icon            =   "PwdChg.frx":0000
   Left            =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Top             =   1110
   Width           =   5490
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      Index           =   1
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtPwd 
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Caption         =   "Con&firm new password:"
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblLabel 
      Caption         =   "&New password:"
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Caption         =   "&Old password:"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfirmChangePwd"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmConfirmChangePwd
'***************************************************************************


Private Sub cmdOKCancel_Click(Index As Integer)
'EncryptPassword
  Select Case Index
    Case 0
      If txtPwd(1) <> txtPwd(2) Then
        MsgBox "The password you type is not match"
        txtPwd(2).SetFocus
        Exit Sub
      ElseIf txtPwd(2) = "admin" Then
      ' The password is to obvious; NEXT VERSION create common password list for restriction
        MsgBox "That password is not allowed"
        txtPwd(2).SetFocus
        Exit Sub
      End If
      
      If boolSetPwdLen = True Then
        If Len(txtPwd(2)) < intSetPwdLen Then
          MsgBox "Your password must be at least " & intSetPwdLen & " characters long", vbExclamation + vbOKOnly, strSystem
          txtPwd(2).SetFocus
          Exit Sub
        End If
      End If
      
      If boolSetPwdAlphaNum = True Then
        If chkIsAlphaNum(CStr(txtPwd(2))) = False Then
          MsgBox "Your password must contain a combination of letters and numbers", vbExclamation + vbOKOnly, strSystem
          txtPwd(2).SetFocus
          Exit Sub
        End If
      End If
      
        strPasswd$ = txtPwd(2)
        SaveSetting strExecName, "User name", currUsr$, txtPwd(2) ' password encryption removed
        MsgBox "Password changed succesfully."
        Unload Me
    Case 1
      MsgBox "Password doesn't changed."
      Unload Me
  End Select
  End Sub
Private Sub Form_Load()
  Caption = "Change " & strSystem & " Password"
  txtPwd(0) = strPasswd$
  txtPwd(0).Enabled = False
End Sub
Private Sub txtPwd_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index < 2 Then
    KeyAscii = 0
    txtPwd(Index + 1).SetFocus
  End If
  If KeyAscii = 13 And Index = 2 Then
    KeyAscii = 0
    cmdOKCancel(0).SetFocus
  End If
End Sub
