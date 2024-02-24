VERSION 4.00
Begin VB.Form frmLibOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Options"
   ClientHeight    =   3840
   ClientLeft      =   2655
   ClientTop       =   2430
   ClientWidth     =   6555
   Height          =   4245
   Icon            =   "Options.frx":0000
   Left            =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Top             =   2085
   Width           =   6675
   Begin VB.Frame fraDBUtil 
      Caption         =   "Database Utilities"
      Height          =   1575
      Left            =   3480
      TabIndex        =   10
      Top             =   1800
      Width           =   3015
      Begin Threed.SSCommand cmd3DBUtil 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Repair ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmd3DBUtil 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "New ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmd3DBUtil 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Compact ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Outline         =   0   'False
      End
   End
   Begin VB.Frame fraILISOptions 
      Caption         =   "Inpens Library Information System Options"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6375
      Begin VB.OptionButton optFineRate 
         Caption         =   "Set Fine Per MemberType"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   3255
      End
      Begin VB.OptionButton optFineRate 
         Caption         =   "Set Fine Per Item Type"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.ComboBox cboMemberType 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   8
         Text            =   "7"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   20
         Top             =   1350
         Width           =   4575
      End
      Begin VB.Label lblInfo 
         Caption         =   "Member Type:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblInfo 
         Caption         =   "Set Fine Rate"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label iblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Days for claimants to claim the items  :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame fraLvwOptions 
      Caption         =   "List View Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
      Begin VB.CheckBox chkShowLvw 
         Caption         =   "Select the whole row when clicked"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton OptShowAllLvw 
         Caption         =   "Show all contents"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton OptShowAllLvw 
         Caption         =   "Show all headers"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fraLibDB 
      Caption         =   "Set default database"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5535
      End
      Begin Threed.SSCommand cmd3Browse 
         Height          =   300
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Outline         =   0   'False
      End
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   5280
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin MSComDlg.CommonDialog cmlLibrary 
      Left            =   960
      Top             =   5280
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLibOptions"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cboMemberType_Change()
    cboMemberType_Click
End Sub


Private Sub cboMemberType_Click()
    SQL = "SELECT [MEMBER-TYPE].MEMBER_DESC FROM [MEMBER-TYPE] WHERE MEMBER_CODE = '" & cboMemberType & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    If Not gLibSnap.EOF Then
        lblInfo(2) = gLibSnap.Fields(0)
    End If
End Sub


Private Sub chkShowLvw_Click()
    saveLibSet "SelectRow", CStr(chkShowLvw.Value)
End Sub


Private Sub cmd3Browse_Click()
    With frmLibOptions.cmlLibrary
        .Filter = "Library Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
        .FilterIndex = 0
        .Flags = FileOpenConstants.cdlOFNHideReadOnly
        .ShowOpen
        If .FileName = "*.mdb" Then
            Exit Sub
        ElseIf .FileName = "" Then
            MsgBox "Changes unsuccessful.", vbInformation, strSystem
            Exit Sub
        End If
        
        If MsgBox("Do you want to specify the database " & .FileName & _
            " to default database.", vbYesNo + vbQuestion + _
            vbDefaultButton1, strSystem) = vbYes Then
            SaveSetting strExecName, "Settings", "Database", .FileName
            txtinfo(1) = .FileName
            MsgBox "Please restart " & strSystem & " to take the changes take effect.", vbInformation, strSystem
        End If
    End With
End Sub


Private Sub cmd3Close_Click()
    saveLibSet "ClaimDay", txtinfo(0).Text
    Unload Me
End Sub


Private Sub cmd3DBUtil_Click(Index As Integer)
    Select Case Index
        Case 0
            CreateNewDB
        Case 2
            DBRepair
    End Select
End Sub


Private Sub Form_Load()
    txtinfo(0) = getLibSet("ClaimDay")
    txtinfo(1) = getLibSet("Database")
    chkShowLvw.Value = CInt(getLibSet("SelectRow"))
    If getLibSet("ShowHeader") = "False" Then
        OptShowAllLvw(0).Value = 1
    ElseIf getLibSet("ShowHeader") = "True" Then
        OptShowAllLvw(1).Value = 1
    End If
    dbRetrieve cboMemberType, "MEMBER-TYPE", "MEMBER_CODE", 0
End Sub


Private Sub OptShowAllLvw_Click(Index As Integer)
    Select Case Index
        Case 0
            saveLibSet "ShowHeader", "False"
        Case 1
            saveLibSet "ShowHeader", "True"
    End Select
End Sub
