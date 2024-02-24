VERSION 4.00
Begin VB.Form frmSubject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject"
   ClientHeight    =   4560
   ClientLeft      =   2100
   ClientTop       =   960
   ClientWidth     =   6210
   Height          =   4965
   Icon            =   "LSubject.frx":0000
   Left            =   2040
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6210
   Top             =   615
   Width           =   6330
   Begin VB.Frame fraSubject 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   6015
      Begin VB.TextBox txtFields 
         DataField       =   "SUBJECT_CODE"
         DataSource      =   "datSubject"
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SUBJECT_DESC"
         DataSource      =   "datSubject"
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   1080
         Width           =   4095
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   8
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
         Left            =   2400
         TabIndex        =   7
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
         Left            =   3480
         TabIndex        =   6
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
         Caption         =   "Subject Code"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Caption         =   "Subject Description"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
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
      TabIndex        =   11
      Top             =   4200
      Width           =   4455
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   5160
      TabIndex        =   10
      Top             =   4200
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
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   4260
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
      i1              =   "LSubject.frx":0442
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
      i1              =   "LSubject.frx":0939
   End
End
Attribute VB_Name = "frmSubject"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmSubject
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
    gtTableActive = "SUBJECT"
    gtTableIndex = "SUBJECT_CODE"
    
    gtMainLvw = "RecSubject"
    
    ShowFirst Me
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
    If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtFields(Index + 1).SetFocus
        Case 1
            cmd3DBCtrl(2).SetFocus
    End Select
    KeyAscii = 0
    End If
End Sub
