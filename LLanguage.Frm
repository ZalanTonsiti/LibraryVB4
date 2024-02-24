VERSION 4.00
Begin VB.Form frmLanguage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language"
   ClientHeight    =   4680
   ClientLeft      =   1980
   ClientTop       =   1890
   ClientWidth     =   6240
   Height          =   5085
   Left            =   1920
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Top             =   1545
   Width           =   6360
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   6015
      Begin VB.TextBox txtFields 
         DataField       =   "LANGUAGE_CODE"
         DataSource      =   "datLanguage"
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtFields 
         DataField       =   "LANGUAGE_DESC"
         DataSource      =   "datLanguage"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   3
         Left            =   3480
         TabIndex        =   9
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
         HelpContextID   =   7
         Index           =   2
         Left            =   2400
         TabIndex        =   8
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
         HelpContextID   =   94
         Index           =   1
         Left            =   1320
         TabIndex        =   7
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
         HelpContextID   =   5
         Index           =   0
         Left            =   240
         TabIndex        =   6
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
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Language Code"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Language Description"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Label lblIndicator 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   4575
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   5160
      TabIndex        =   10
      Top             =   4320
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
   Begin ComctlLib.ListView lvwLibrary 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   4895
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
      SmallIcons      =   "imlSmallIcon"
      Sorted          =   -1  'True
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
      i1              =   "LLanguage.frx":0000
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
      i1              =   "LLanguage.frx":04F7
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmLanguage
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
    gtTableActive = "LANGUAGE"
    gtTableIndex = "LANGUAGE_CODE"
    gtMainLvw = "RecLanguage"
    ShowFirst frmLanguage
    fillLvw Me, 2
    FieldControlDisabled Me
    CtrEnableLvwViewMain
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gtTableIndex = ""
    CtrEnableLvwViewMain
    gtMainLvw = "MainCataloging"
End Sub


Private Sub lvwLibrary_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  lvwLibrary.SortKey = ColumnHeader.Index - 1
  lvwLibrary.SortOrder = Abs(CInt(gfSort))      ' Ascending
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
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    lblIndicator = lblFields(Index) & " ..."
End Sub
