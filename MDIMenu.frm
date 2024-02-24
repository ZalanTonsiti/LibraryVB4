VERSION 4.00
Begin VB.MDIForm frmMDIMainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "LIBRARY INFORMATION SYSTEM 1.0"
   ClientHeight    =   4395
   ClientLeft      =   1470
   ClientTop       =   2670
   ClientWidth     =   9825
   Height          =   5085
   Icon            =   "MDIMenu.frx":0000
   Left            =   1410
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Top             =   2040
   Width           =   9945
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrLibrary 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   688
      AllowCustomize  =   0   'False
      ImageList       =   "imlLibrary"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "New"
            ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Del"
            ToolTipText     =   "Delete selected Record"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            ToolTipText     =   "Find"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Properties"
            ToolTipText     =   "Properties"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "lvwView0"
            ToolTipText     =   "Large Icons"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "lvwView1"
            ToolTipText     =   "Small Icons"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "lvwView2"
            ToolTipText     =   "Lists"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "lvwView3"
            ToolTipText     =   "Details"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "Help"
            ToolTipText     =   "Help"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmlLibrary 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport repMainLibrary 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ReportFileName  =   ""
      Destination     =   0
      WindowLeft      =   100
      WindowTop       =   100
      WindowWidth     =   490
      WindowHeight    =   300
      WindowTitle     =   ""
      WindowBorderStyle=   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   1
      PrintFileName   =   ""
      PrintFileType   =   0
      SelectionFormula=   ""
      GroupSelectionFormula=   ""
      Connect         =   ""
      UserName        =   ""
      ReportSource    =   0
      BoundReportHeading=   ""
      BoundReportFooter=   0   'False
   End
   Begin ComctlLib.ImageList imlLibrary 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":06EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0800
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0912
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMenu.frx":0E6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sbrLibrary 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   4155
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14261
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Library 
      Caption         =   "&Library"
      Begin VB.Menu Circulation 
         Caption         =   "&Circulation"
         Begin VB.Menu mnuCirculation 
            Caption         =   "&Loan                      "
            Index           =   0
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuCirculation 
            Caption         =   "Ret&urn                  "
            Index           =   1
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuCirculation 
            Caption         =   "L&ost Material          "
            Index           =   2
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu Cataloging 
         Caption         =   "Catalogin&g"
         Begin VB.Menu mnuCataloging 
            Caption         =   "Mater&ial"
            Index           =   0
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuCataloging 
            Caption         =   "Mat&erial Type"
            Index           =   1
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuCataloging 
            Caption         =   "&Placement"
            Index           =   2
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuCataloging 
            Caption         =   "&Subject"
            Index           =   3
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuCataloging 
            Caption         =   "Lan&guage"
            Index           =   4
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu Acquisition 
         Caption         =   "&Acquisition"
         Begin VB.Menu mnuAcquisition 
            Caption         =   "Pu&blisher"
            Index           =   0
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuAcquisition 
            Caption         =   "&Vendor"
            Index           =   1
            Shortcut        =   ^V
         End
         Begin VB.Menu mnuAcquisition 
            Caption         =   "Suppl&y"
            Index           =   2
            Shortcut        =   ^Y
         End
      End
      Begin VB.Menu Patron 
         Caption         =   "&Patron"
         Begin VB.Menu mnuPatron 
            Caption         =   "&Member"
            Index           =   0
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuPatron 
            Caption         =   "Member &Type"
            Index           =   1
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu Utility 
         Caption         =   "&Utility"
         Begin VB.Menu mnuUtility 
            Caption         =   "Boo&king"
            Index           =   0
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuUtility 
            Caption         =   "Use&r"
            Index           =   1
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuUtility 
            Caption         =   "A&nnual Update"
            Index           =   2
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu Separator4 
         Caption         =   "-"
      End
      Begin VB.Menu Report 
         Caption         =   "&Report"
      End
      Begin VB.Menu Separator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "&Log Off"
         Index           =   0
      End
      Begin VB.Menu mnuLibrary 
         Caption         =   "Log Off and E&xit"
         Index           =   1
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRec 
         Caption         =   "Edit"
      End
      Begin VB.Menu Separator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu View 
      Caption         =   "&View"
      Begin VB.Menu mnuViewElip 
         Caption         =   "Lar&ge Icon"
         Index           =   0
      End
      Begin VB.Menu mnuViewElip 
         Caption         =   "S&mall Icon"
         Index           =   1
      End
      Begin VB.Menu mnuViewElip 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuViewElip 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu Separator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort"
         Begin VB.Menu mnuSortElip 
            Caption         =   "&Ascending"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuSortElip 
            Caption         =   "&Descending"
            Index           =   1
         End
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "&Tool"
      Begin VB.Menu mnuFind 
         Caption         =   "Find..."
         Visible         =   0   'False
      End
      Begin VB.Menu Separator9 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSec 
         Caption         =   "Change Password"
      End
      Begin VB.Menu Separator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelp 
         Caption         =   "Contents"
         Index           =   0
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "About..."
         Index           =   2
      End
   End
   Begin VB.Menu Graph 
      Caption         =   "Graph"
      Visible         =   0   'False
      Begin VB.Menu mnuGMat_Lang_In 
         Caption         =   "Material by Language (Inventory)"
      End
      Begin VB.Menu mnuGMember_MemType 
         Caption         =   "Member by Member Type"
      End
   End
End
Attribute VB_Name = "frmMDIMainMenu"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMDIMainMenu
'***************************************************************************

Private libSnap As Recordset
Private pfFindButtonState As Boolean
Private Declare Function ShellAbout Lib "shell32.dll" _
Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Sub MDIForm_Load()
    CenterForm Me
    frmContent.Show
    DoEvents
    frmContent.ZOrder 0
    frmPwd.Show 1
    pfFindButtonState = tbrLibrary.Buttons("Find").Value
    CtrDisableLvwView
End Sub


Private Sub MDIForm_Resize()
    If WindowState <> 1 Then
        WindowState = 2
    frmContent.Left = 0
    frmContent.Top = 0
    frmContent.Width = 11930
    frmContent.Height = frmMDIMainMenu.Height - 1380
    End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    ' Ask user for confirmation
    If MsgBox("Quit Application?", vbQuestion + vbYesNo + vbDefaultButton2, _
        strSystem) = vbNo Then
        ' Don't allow close
        Cancel = -1
        ShowFirstGrp
        frmContent.Show
    Else
        SaveSetting strExecName, "Access", "Last Access", Date
        saveLibSet "Last Access Date", Date
        End
    End If
End Sub


Private Sub mnuAcquisition_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPublisher.Show
        Case 1
            frmVendor.Show
        Case 2
            frmSupply.Show
    End Select
End Sub


Private Sub mnuCataloging_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMaterial.Show
        Case 1
            frmMaterialType.Show
        Case 2
            frmPlacement.Show
        Case 3
            frmSubject.Show
        Case 4
            frmLanguage.Show
    End Select
End Sub


Private Sub mnuCirculation_Click(Index As Integer)
    Select Case Index
        Case 0
            frmLoan.Show
        Case 1
            frmReturn.Show
        Case 2
            frmReturnPatronNum.Show
    End Select
End Sub


Public Sub mnuDelete_Click()
    On Local Error GoTo Err_Handler
    Dim ctrFrm As Form
    If gtTableActive = "" Then Exit Sub
    If gtCurrentIndex = "" Then
        MsgBox "Please click a record.", vbOKOnly, strSystem
        Exit Sub
    End If
    If currUsr <> "Admin" Then
        MsgBox "You must log as Admin to delete a record.", vbExclamation _
            + vbOKOnly, strSystem
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete record - " _
        & gtCurrentIndex, vbInformation + vbYesNo + vbDefaultButton2, _
        strSystem) = vbNo Then Exit Sub
    If gtTableActive <> "MATERIAL" Then
        SQL = "SELECT * FROM [" & gtTableActive & "] WHERE [" _
            & gtTableIndex & "] = '" & gtCurrentIndex & "'"
    Else
        SQL = "SELECT * FROM [" & gtTableActive & "] WHERE [" _
            & gtTableIndex & "] = " & gtCurrentIndex
    End If
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    With gLibDyna
        While Not .EOF
            .Delete
            .MoveNext
        Wend
    End With
    Select Case gtMainLvw
        Case "MainAcquisition"
            Set ctrFrm = frmMainAcquisition
        Case "MainCataloging"
            Set ctrFrm = frmMainCataloging
        Case "MainCirculation"
            Set ctrFrm = frmMainCirculation
        Case "MainPatron"
            Set ctrFrm = frmMainPatron
        Case "MainUtility"
            Set ctrFrm = frmMainUtility
        Case "RecPublisher"
            Set ctrFrm = frmPublisher
        Case "RecVendor"
            Set ctrFrm = frmVendor
        Case "RecSupply"
            Set ctrFrm = frmSupply
        Case "RecMaterial"
            Set ctrFrm = frmMaterial
        Case "RecMaterialType"
            Set ctrFrm = frmMaterialType
        Case "RecPlacement"
            Set ctrFrm = frmPlacement
        Case "RecSubject"
            Set ctrFrm = frmSubject
        Case "RecLanguage"
            Set ctrFrm = frmLanguage
'        Case "RecLoan"
'            Set ctrFrm = frmLoan
'        Case "RecReturn"
'            Set ctrFrm = frmReturn
        Case "RecMember"
            Set ctrFrm = frmMember
        Case "RecMemberType"
            Set ctrFrm = frmMemberType
    End Select
    fillLvw ctrFrm, 4
    GoTo End_Sub
Err_Handler:
    DisplayError gtLibErr(25)
    Resume End_Sub
End_Sub:
End Sub


Public Sub mnuGMat_Lang_In_Click()
On Local Error GoTo Error_Handler
    frmContent.grplibrary.YAxisStyle = 2
    frmContent.grplibrary.YAxisMax = 400
    frmContent.grplibrary.YAxisMin = 0
    frmContent.grplibrary.YAxisTicks = 4
    Dim iRecCount As Integer
    Dim iIndex As Integer
    frmContent.grplibrary.GraphTitle = "Material by Language"
    frmContent.grplibrary.LeftTitle = "Material"
    frmContent.grplibrary.BottomTitle = "Language"
    Set gLibSnap = gLibDB.OpenRecordset("LANGUAGE")
    
 If Not gLibSnap.EOF Then
    iRecCount = gLibSnap.RecordCount
    frmContent.grplibrary.NumPoints = iRecCount
    SQL = "SELECT LANGUAGE_CODE, LANGUAGE_DESC FROM LANGUAGE"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
     For iIndex = 1 To iRecCount
        frmContent.grplibrary.ThisPoint = iIndex
        SQL = "SELECT * FROM MATERIAL WHERE LANGUAGE_CODE = '" _
            & gLibSnap.Fields(0) & "'"
        Set libSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If Not libSnap.EOF Then
            libSnap.MoveLast
            frmContent.grplibrary.GraphData = CInt(libSnap.RecordCount)
        ElseIf libSnap.EOF Then
            frmContent.grplibrary.GraphData = 0
        End If
        frmContent.grplibrary.NumSets = 1
        frmContent.grplibrary.ThisPoint = iIndex
        frmContent.grplibrary.LabelText = CStr(gLibSnap.Fields(0))
        frmContent.grplibrary.ThisPoint = iIndex
        'CStr(gLibSnap.Fields(0)) & ", " &
        frmContent.grplibrary.LegendText = CStr(gLibSnap.Fields(1)) & ", " & CStr(libSnap.RecordCount) '& " unit(s)"
        gLibSnap.MoveNext
    Next
  End If
    frmContent.grplibrary.DrawMode = 2
    GoTo End_Sub
Error_Handler:
DisplayError ""
Resume End_Sub
End_Sub:
End Sub


Private Sub mnuGMember_MemType_Click()
On Local Error GoTo Error_Handler
    frmContent.grplibrary.YAxisStyle = 2
    frmContent.grplibrary.YAxisMax = 150
    frmContent.grplibrary.YAxisMin = 0
    frmContent.grplibrary.YAxisTicks = 15
    Dim iRecCount As Integer
    Dim iIndex As Integer
    frmContent.grplibrary.GraphTitle = "Member by Member Type"
    frmContent.grplibrary.LeftTitle = "Member"
    frmContent.grplibrary.BottomTitle = "Member Type"
    Set gLibSnap = gLibDB.OpenRecordset("Member-Type")
  If Not gLibSnap.EOF Then
    iRecCount = gLibSnap.RecordCount
    frmContent.grplibrary.NumPoints = iRecCount
    SQL = "SELECT MEMBER_CODE, MEMBER_DESC FROM [MEMBER-TYPE]"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
     For iIndex = 1 To iRecCount
        frmContent.grplibrary.ThisPoint = iIndex
        SQL = "SELECT * FROM PATRON WHERE MEMBER_CODE = '" _
            & gLibSnap.Fields(0) & "'"
        Set libSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If Not libSnap.EOF Then
            libSnap.MoveLast
            frmContent.grplibrary.GraphData = CInt(libSnap.RecordCount)
        ElseIf libSnap.EOF Then
            frmContent.grplibrary.GraphData = 0
        End If
        frmContent.grplibrary.NumSets = 1
        frmContent.grplibrary.ThisPoint = iIndex
        frmContent.grplibrary.LabelText = CStr(gLibSnap.Fields(0))
        frmContent.grplibrary.ThisPoint = iIndex
        frmContent.grplibrary.LegendText = CStr(gLibSnap.Fields(0)) _
            & ", " & CStr(gLibSnap.Fields(1))
        gLibSnap.MoveNext
    Next
  End If
    frmContent.grplibrary.DrawMode = 2
    GoTo End_Sub
Error_Handler:
DisplayError ""
Resume End_Sub
End_Sub:

End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
        Case 0
            With cmlLibrary
                .HelpCommand = cdlHelpContents
                .HelpFile = GetPath & "Lib.hlp"
                .ShowHelp
            End With
        Case 2
            frmAboutILIS.Show 1
    End Select
End Sub


Private Sub mnuLibrary_Click(Index As Integer)
    Select Case Index
        Case 0
            gfLogConfirm = False ' Reinit Pwd Confirmation
            frmContent.piclibrary.Visible = False
            frmContent.grplibrary.Visible = False
            frmContent.imgLogo.Visible = True
            ' Bring Content in front regardless current form shown
            frmContent.ZOrder 0
            frmPwd.Show 1
        Case 1
            MDIForm_Unload (0)
    End Select
End Sub


Private Sub mnuOptions_Click()
    If currUsr <> "Admin" Then
        MsgBox "You need Admin access", vbExclamation + vbOKOnly, strSystem
        Exit Sub
    Else
        frmLibOptions.Show 1
    End If
End Sub

Private Sub mnuPatron_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMember.Show
        Case 1
            frmMemberType.Show
    End Select
End Sub


Private Sub mnuReport_Click(Index As Integer)
    Select Case Index
        Case 0
            frmGenReport.Show
        Case 1
    End Select
End Sub


Private Sub mnuSec_Click()
    frmConfirmChangePwd.Show vbModal
End Sub


Private Sub mnuSortElip_Click(Index As Integer)
  mnuSortElip(0).Checked = Not mnuSortElip(0).Checked
  mnuSortElip(1).Checked = Not mnuSortElip(1).Checked
  gfSort = Not gfSort
End Sub

Public Sub mnuUtility_Click(Index As Integer)
    Select Case Index
        Case 0
            frmBooking.Show
        Case 1                  ' User
            frmAdminPwd.Show 1
        Case 2                  ' Annual update
            dbAnnualUpdate
    End Select
End Sub


Private Sub mnuViewElip_Click(Index As Integer)
  changeViewMain ("lvwView" & CStr(Index))
End Sub


Private Sub Report_Click()
    frmGenReport.Show
End Sub

Private Sub tbrLibrary_ButtonClick(ByVal Button As Button)
    Dim ctrFrm As Form
  Select Case Button.Key
    Case "New"
        Set ctrFrm = ctrGetCurrentForm
        ctrFrm.cmd3DBCtrl_Click (0)
    Case "Del"
        mnuDelete_Click
    Case "Properties"
        gCtrFrm.Show
    Case "Find"
        If pfFindButtonState Then
            tbrLibrary.Buttons("Find").Value = tbrUnpressed
            ctrFormSeacrhtoRec
            gfSearchMode = False
            Set ctrFrm = ctrGetCurrentForm
            ctrFrm.Form_Load
                If gtMainLvw = "RecMaterial" Then
                   ctrFrm.lstSubject.Enabled = False
                   ctrFrm.cboExistence.Enabled = False
                End If
        ElseIf Not pfFindButtonState Then
            tbrLibrary.Buttons("Find").Value = tbrPressed
            ctrFormRectoSeacrh
            gfSearchMode = True
            Set ctrFrm = ctrGetCurrentForm
            clearLvw ctrFrm.lvwLibrary
                If gtMainLvw = "RecMaterial" Then
                    ctrFrm.lstSubject.Clear
                    ctrFrm.lstSubject.AddItem ""
                    ctrFrm.lstSubject.Enabled = True
                    ctrFrm.cboExistence.Enabled = True
                End If
        End If
        pfFindButtonState = tbrLibrary.Buttons("Find").Value
    Case "Del"
        mnuDelete_Click
    Case "lvwView0", "lvwView1", "lvwView2", "lvwView3"
      If gtMainLvw <> "" Then
          ' Change for the Main*.ListView.View
          changeViewMain (Button.Key)
      Else
      End If
    Case "Print"
        Select Case gtMainLvw
            Case "RecSubject"
                With repmainlibrary
                    .Destination = 0
                    .DataFiles(0) = getLibSet("Database")
                    .ReportFileName = GetPath & "List of Subject.rpt"
                    .Action = 1
                End With
            
            Case "RecMaterialType"
                With repmainlibrary
                    .Destination = 0
                    .DataFiles(0) = getLibSet("Database")
                    .ReportFileName = GetPath & "List of Material Type.rpt"
                    .Action = 1
                End With
            
            Case "RecPlacement"
                With repmainlibrary
                    .Destination = 0
                    .DataFiles(0) = getLibSet("Database")
                    .ReportFileName = GetPath & "List of Placements.rpt"
                    .Action = 1
                End With
        
            Case "RecLanguage"
                With repmainlibrary
                    .Destination = 0
                    .DataFiles(0) = getLibSet("Database")
                    .ReportFileName = GetPath & "List of Languages.rpt"
                    .Action = 1
                End With
            
            Case "RecPatron"
                frmGetPatronNum.Show 1
                With repmainlibrary
                    .SelectionFormula = "{PATRON.PATRON_NUM} = '" & gtGetPatron_Num & "'"
                    .Destination = 0
                    .DataFiles(0) = getLibSet("Database")
                    .ReportFileName = GetPath & "List of Members.rpt"
                    .Action = 1
                End With
            Case "GenReport"
                frmGenReport.cmd3GenReport_Click (0)
        End Select
        Case "Help"
            mnuHelp_Click (0)
  End Select
End Sub


