VERSION 4.00
Begin VB.Form frmVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor"
   ClientHeight    =   5670
   ClientLeft      =   1590
   ClientTop       =   1170
   ClientWidth     =   7080
   Height          =   6075
   Icon            =   "LVendor.frx":0000
   Left            =   1530
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7080
   Top             =   825
   Width           =   7200
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   6855
      Begin VB.TextBox txtFields 
         DataField       =   "VENDOR_CODE"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "VENDOR_NAME"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   930
         Width           =   5055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ADDRESS"
         DataSource      =   "datVendor"
         Height          =   615
         Index           =   2
         Left            =   1680
         TabIndex        =   5
         Top             =   1260
         Width           =   5055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "FAX_NUMBER"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   4
         Text            =   "019-1234567"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         DataField       =   "TELEPHONE_OFF"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   3
         Text            =   "019-1234567"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         DataField       =   "TELEPHONE_HP"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   6
         Left            =   5640
         TabIndex        =   2
         Text            =   "019-1234567"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CONTACT_PERSON"
         DataSource      =   "datVendor"
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   1
         Top             =   1920
         Width           =   5055
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   19
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
         Left            =   1320
         TabIndex        =   18
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
         Left            =   2880
         TabIndex        =   17
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
         Left            =   3960
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
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Code"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor Name"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Hand Phone Number"
         Height          =   375
         Index           =   6
         Left            =   4440
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone Number"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax Number"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblFields 
         Caption         =   "Address"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Person"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
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
      TabIndex        =   21
      Top             =   5280
      Width           =   5655
   End
   Begin ComctlLib.ListView lvwLibrary 
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   4048
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
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   6000
      TabIndex        =   20
      Top             =   5280
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
      i1              =   "LVendor.frx":0442
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
      i1              =   "LVendor.frx":0939
   End
End
Attribute VB_Name = "frmVendor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmVendor
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
    gtTableActive = "VENDOR"
    gtTableIndex = "VENDOR_CODE"
    
    gtMainLvw = "RecVendor"
    
    ShowFirst Me
    fillLvw Me, 5
    FieldControlDisabled Me
    CtrEnableLvwViewMain
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gtTableIndex = ""
    CtrEnableLvwViewMain
    gtMainLvw = "MainAcquisition"
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
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    lblIndicator = lblFields(Index) & " ..."
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And Index < 7 Then
    txtFields(Index + 1).SetFocus
End If
End Sub
