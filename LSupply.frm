VERSION 4.00
Begin VB.Form frmSupply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supply"
   ClientHeight    =   4680
   ClientLeft      =   1920
   ClientTop       =   2265
   ClientWidth     =   6210
   Height          =   5085
   Icon            =   "LSupply.frx":0000
   Left            =   1860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6210
   Top             =   1920
   Width           =   6330
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   6015
      Begin VB.TextBox txtFields 
         DataField       =   "ACQUISITION_NUM"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   9
         Tag             =   "Acquisition No"
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Tag             =   "Vendor Code"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Height          =   300
         Index           =   2
         Left            =   4800
         TabIndex        =   11
         Tag             =   "Supply Date"
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Height          =   300
         Index           =   3
         Left            =   4800
         TabIndex        =   13
         Tag             =   "Price"
         Text            =   "0"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Price"
         Height          =   225
         Index           =   3
         Left            =   4200
         TabIndex        =   14
         Tag             =   "Price"
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Supply Date"
         Height          =   225
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Tag             =   "Supply Date"
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition No"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Code"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   0
         Left            =   120
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
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   5
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
         Index           =   2
         Left            =   2280
         TabIndex        =   4
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
         Index           =   3
         Left            =   3360
         TabIndex        =   3
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
      i1              =   "LSupply.frx":0442
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
      i1              =   "LSupply.frx":0939
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
      TabIndex        =   16
      Top             =   4320
      Width           =   4935
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   5160
      TabIndex        =   2
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   4683
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
   End
End
Attribute VB_Name = "frmSupply"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmSupply
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
    gtTableActive = "SUPPLY"
    gtTableIndex = "ACQUISITION_NUM"
    
    gtMainLvw = "RecSupply"
    
    ShowFirst frmSupply
    fillLvw Me, 4
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
  lvwLibrary.SortOrder = 0       ' Ascending
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
    With frmMDIMainMenu.tbrLibrary
    Select Case Index
        Case 0
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmMaterial
        Case 1
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmVendor
        Case Else
            .Buttons("Properties").Enabled = False
    End Select
    End With
End Sub
