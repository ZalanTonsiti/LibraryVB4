VERSION 4.00
Begin VB.Form frmAdminPwd 
   Caption         =   "Adminstration"
   ClientHeight    =   2535
   ClientLeft      =   690
   ClientTop       =   3510
   ClientWidth     =   8070
   Height          =   3225
   HelpContextID   =   10
   Icon            =   "PwdAdmin.frx":0000
   Left            =   630
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   8070
   Top             =   2880
   Width           =   8190
   Begin ComctlLib.ImageList ImageList2 
      Left            =   2520
      Top             =   4560
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      NumImages       =   1
      i1              =   "PwdAdmin.frx":000C
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   4560
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      NumImages       =   1
      i1              =   "PwdAdmin.frx":0503
   End
   Begin ComctlLib.ListView lvwPasswd 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   4471
      _StockProps     =   205
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Icons           =   ""
      LabelEdit       =   1
      SmallIcons      =   ""
   End
   Begin VB.Menu mnuPasswd 
      Caption         =   "&User"
      Begin VB.Menu mnuUser 
         Caption         =   "&Add"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&Edit"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&Delete"
         Index           =   2
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPasswd2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu Tool 
         Caption         =   "&Find"
         Index           =   0
      End
      Begin VB.Menu Tool 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Tool 
         Caption         =   "&Options..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuView_ 
      Caption         =   "&View"
      Begin VB.Menu View 
         Caption         =   "&Small Icons"
         Index           =   0
      End
      Begin VB.Menu View 
         Caption         =   "Large &Icons"
         Index           =   1
      End
      Begin VB.Menu View 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu View 
         Caption         =   "De&tails"
         Index           =   3
      End
      Begin VB.Menu View 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort"
         Begin VB.Menu Sort 
            Caption         =   "Ascending"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu Sort 
            Caption         =   "Descending"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPasswd3 
         Caption         =   "-"
      End
      Begin VB.Menu View2 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmAdminPwd"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmAdminPwd
'***************************************************************************

Public selUsr               As String            ' Selected user
Public boolSort             As Boolean
Public intTotalUsr          As Integer           ' Total User

Sub clrView()
  lvwPasswd.View = lvwReport
  lvwPasswd.ListItems.Clear
End Sub

Sub fillColHeader()
  Dim clmX As ColumnHeader
  Dim imgX As ListImage
    lvwPasswd.Icons = ImageList1
    lvwPasswd.SmallIcons = ImageList2
  Set clmX = lvwPasswd.ColumnHeaders.Add(, , "User name", lvwPasswd.Width / 8)
  Set clmX = lvwPasswd.ColumnHeaders.Add(, , "Password", lvwPasswd.Width / 8)
  Set clmX = lvwPasswd.ColumnHeaders.Add(, , "Name", lvwPasswd.Width / 2)
  Set clmX = lvwPasswd.ColumnHeaders.Add(, , "Created", lvwPasswd.Width / 4)
  Set clmX = lvwPasswd.ColumnHeaders.Add(, , "Last Log", lvwPasswd.Width / 4)
End Sub

Private Sub Form_GotFocus()
  updateGetAllSettings
End Sub

Private Sub Form_Load()
  fillColHeader
  updateGetAllSettings
End Sub

Private Sub Form_Resize()
  If frmAdminPwd.Height > 1200 Then lvwPasswd.Height = frmAdminPwd.Height - 690
  If frmAdminPwd.Width > 1000 Then lvwPasswd.Width = frmAdminPwd.Width - 120
End Sub

Private Sub lvwPasswd_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  lvwPasswd.SortKey = ColumnHeader.Index - 1
  lvwPasswd.SortOrder = Abs(CInt(boolSort))
  lvwPasswd.Sorted = 1
End Sub

Private Sub lvwPasswd_LostFocus()
  Dim i As Integer
  For i = 1 To lvwPasswd.ListItems.Count
    lvwPasswd.ListItems.Item(i).Selected = False
  Next i
End Sub

Private Sub lvwPasswd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbRightButton Then PopupMenu mnuPasswd
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub mnuUser_Click(Index As Integer)
On Local Error GoTo Error_Handler
Select Case Index
  Case 0
    If currUsr <> "Admin" Then
      MsgBox "You need Admin access", vbExclamation + vbOKOnly, strSystem
      GoTo End_Sub
    ElseIf currUsr = "Admin" Then
      ' Show Pwd Dialog Box again for confirmation
      frmPwd.Show 1
    End If
    'frmAddState = True
'           SUPPOSE NOT TO BE HERE    frmAddEditUsr.Show vbModal
            '  Case 1 ' Menu is hidden FOR NEXT VERSION
            '    'frmAddState = False
            '    frmAddEditUsr.Show vbModal
  Case 2
    If currUsr <> "Admin" Then
      MsgBox "You need Supervisor access", vbExclamation + vbOKOnly, strSystem
      GoTo End_Sub
    Else
        On Error GoTo Error_Handler
        selUsr$ = lvwPasswd.SelectedItem
        If MsgBox("Are you sure you want to delete this user?", vbQuestion + vbYesNo, strSystem) = vbYes Then
        ' If selUsr$ = "Admin" Then MsgBox "You can't delete Admin!", vbExclamation, "Error"
        ' If selUsr$ <> "Admin" Then delUsr
          delUsr
        End If
    End If
  End Select
  Exit Sub
Error_Handler:
    DisplayError ""
End_Sub:
End Sub

Private Sub Sort_Click(Index As Integer)
  Index = 0
  Sort(Index).Checked = Not Sort(Index).Checked
  Sort(Index + 1).Checked = Not Sort(Index + 1).Checked
  boolSort = Not boolSort
End Sub

Private Sub Tool_Click(Index As Integer)
  Select Case Index
    Case 0
      Dim strFindMe As String
      strFindMe = InputBox("Find : ")
      Dim itmFound As ListItem  ' FoundItem variable.
      Set itmFound = lvwPasswd.FindItem(strFindMe, 0, , lvwPartial)
      
      If itmFound Is Nothing Then  ' If no match, inform user and exit.
        MsgBox "No match found"
        Exit Sub
      Else
        itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
        itmFound.Selected = True ' Select the ListItem.
        lvwPasswd.SetFocus
      End If
      
    Case 2
      frmOptions.Show vbModal
  End Select
End Sub

Private Sub View_Click(Index As Integer)
  lvwPasswd.View = Index
End Sub

Sub updateGetAllSettings()
  On Error Resume Next
   ' Variant to hold 2-dimensional array returned by GetAllSettings.
  Dim PwdSettings         As Variant
  Dim Settings            As Integer        ' This make the Users liimit to 32767
  Dim itmX                As ListItem
  Dim strFill             As String
  Dim strFill2            As String
  lvwPasswd.View = lvwReport
  clrView
  ' Fill ListView with User name, Password, Name, Created, Last log
      PwdSettings = GetAllSettings(appname:=strExecName, section:="User name")
      For Settings = LBound(PwdSettings, 1) To UBound(PwdSettings, 1)
        strFill = PwdSettings(Settings, 0)
        ' Use this below, the Admin is not shown on lvw
'         Set itmX = lvwPasswd.ListItems.Add(Settings, strFill, strFill, 1) ' object.Add(index, key, text, icon, smallIcon)
        ' This instead, WILL include Admin on the lvw
         Set itmX = lvwPasswd.ListItems.Add(, strFill, strFill, 1)
        itmX.Icon = 1 ' Set an icon from ImageList1.
        itmX.SmallIcon = 1  ' Set an icon from ImageList2.
        strFill2 = GetSetting(strExecName, "User name", PwdSettings(Settings, 0), Default:="<none>")
        itmX.SubItems(1) = strFill2
        itmX.SubItems(2) = GetSetting(strExecName, "Name", strFill, Default:="<none>")
        itmX.SubItems(3) = GetSetting(strExecName, "Created", strFill, Default:="<none>")
        itmX.SubItems(4) = GetSetting(strExecName, "Last log", strFill, Default:="<none>")
      Next Settings
      intTotalUsr = Settings    ' Give plus 1 - 'Invinsible' Admin
      AdjustColumnWidth lvwPasswd, False
    '      frmAdminPwd.Caption = "Adminstration with Total Users of " & Settings
End Sub

Public Sub delUsr()
  On Error GoTo Error_Handler
  If selUsr <> "Admin" Then
      lvwPasswd.ListItems.Remove selUsr$
  Else
      MsgBox "Admin cannot be deleted.", vbCritical, strSystem
      GoTo End_Sub
  End If
  DeleteSetting strExecName, "Created", selUsr$
  DeleteSetting strExecName, "Last log", selUsr$
  DeleteSetting strExecName, "Name", selUsr$
  DeleteSetting strExecName, "User name", selUsr$
  GoTo End_Sub
Error_Handler:
    DisplayError ""
Resume End_Sub
End_Sub:
End Sub

Private Sub View2_Click(Index As Integer)
  updateGetAllSettings
End Sub
