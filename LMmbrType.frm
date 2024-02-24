VERSION 4.00
Begin VB.Form frmMemberType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Type"
   ClientHeight    =   4920
   ClientLeft      =   3270
   ClientTop       =   1710
   ClientWidth     =   5850
   Height          =   5325
   Icon            =   "LMmbrType.frx":0000
   Left            =   3210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   5850
   Top             =   1365
   Width           =   5970
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox txtFields 
         DataField       =   "MEMBER_CODE"
         DataSource      =   "datMemberType"
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         DataField       =   "MEMBER_DESC"
         DataSource      =   "datMemberType"
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1035
         Width           =   3735
      End
      Begin VB.TextBox txtFields 
         DataField       =   "MAX_BORROWING"
         DataSource      =   "datMemberType"
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   4
         Top             =   1365
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "BORROWING_DURATION"
         DataSource      =   "datMemberType"
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   2
         Top             =   1695
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "FINE_RATE_PERDAY"
         DataSource      =   "datMemberType"
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   1
         Top             =   2025
         Width           =   1935
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   3
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Enter"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Ed&it"
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
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Ne&w"
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
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Member Code"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   735
         Width           =   1400
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1050
         Width           =   1400
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Borrowing"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1380
         Width           =   1400
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Borrowing Duration"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1710
         Width           =   1400
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fine Rate Per Day"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1400
      End
      Begin VB.Label lblInfo 
         Caption         =   "Days"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Label lblIndicator 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   4455
   End
   Begin ComctlLib.ListView lvwLibrary 
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   3413
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
      Icons           =   "imlLargeIcon"
      LabelEdit       =   1
      SmallIcons      =   "imlSmallIcon"
      Sorted          =   -1  'True
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   4800
      TabIndex        =   3
      Top             =   4560
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
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
   Begin ComctlLib.ImageList imlSmallIcon 
      Left            =   960
      Top             =   0
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      NumImages       =   1
      i1              =   "LMmbrType.frx":0442
   End
   Begin ComctlLib.ImageList imlLargeIcon 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      NumImages       =   1
      i1              =   "LMmbrType.frx":0939
   End
End
Attribute VB_Name = "frmMemberType"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMemberType
'***************************************************************************


Private Sub cmd3Close_Click()
    Unload Me
End Sub


Public Sub cmd3DBCtrl_Click(Index As Integer)
    If gfSearchMode Then
        Select Case Index
            Case 0
                FieldControlNew Me
            Case 2
                BuildSearchSQL Me
        End Select
    ElseIf Not gfSearchMode Then
        Select Case Index
            Case 0
        End Select
        CtrCommand3D Me, Index
    End If
End Sub


Public Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    gtTableActive = "MEMBER-TYPE"
    gtTableIndex = "MEMBER_CODE"
    
    gtMainLvw = "RecMemberType"
    
    ShowFirst frmMemberType
    fillLvw Me, 6
    FieldControlDisabled Me
    CtrEnableLvwViewMain
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gtTableIndex = ""
    CtrEnableLvwViewMain
    gtMainLvw = "MainPatron"
End Sub

Private Sub lvwLibrary_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  lvwLibrary.SortKey = ColumnHeader.Index - 1
  lvwLibrary.SortOrder = Abs(CInt(gfSort))       ' Ascending
  lvwLibrary.Sorted = 1
End Sub


Private Sub lvwLibrary_ItemClick(ByVal Item As ListItem)
  If getLibSet("SelectRow") = "1" Then
    cSelectRowLvw lvwLibrary
  End If
    gtCurrentIndex = CStr(Item)
    fillCurrentDetail Me, CStr(Item)
End Sub


Private Sub txtFields_Change(Index As Integer)
    txtFields(Index).Tag = "Dirty"
    If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtFields(Index + 1).SetFocus
        Case 1
            txtFields(Index + 1).SetFocus
        Case 2
            txtFields(Index + 1).SetFocus
        Case 3
            txtFields(Index + 1).SetFocus
        Case 4
            txtFields(Index + 1).SetFocus
        Case 5
            cmd3DBCtrl(2).SetFocus
    End Select
    KeyAscii = 0
    End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    lblIndicator = lblLabels(Index) & " ..."
End Sub
