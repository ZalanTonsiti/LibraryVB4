VERSION 4.00
Begin VB.Form frmMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Member"
   ClientHeight    =   6615
   ClientLeft      =   1305
   ClientTop       =   1875
   ClientWidth     =   9030
   Height          =   7020
   Icon            =   "LMember.frx":0000
   Left            =   1245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Top             =   1530
   Width           =   9150
   Begin VB.Frame fraMember 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   8775
      Begin VB.TextBox txtFields 
         DataField       =   "PATRON_NUM"
         DataSource      =   "datMember"
         Height          =   330
         Index           =   0
         Left            =   6600
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "MEMBER_CODE"
         DataSource      =   "datMember"
         Height          =   315
         Index           =   1
         Left            =   6600
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         DataField       =   "IC_NUM"
         DataSource      =   "datMember"
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "NAME"
         DataSource      =   "datMember"
         Height          =   330
         Index           =   3
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ADDRESS"
         DataSource      =   "datMember"
         Height          =   990
         Index           =   4
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "TELEPHONE_NUM"
         DataSource      =   "datMember"
         Height          =   330
         Index           =   5
         Left            =   1800
         TabIndex        =   6
         Top             =   2820
         Width           =   2535
      End
      Begin VB.TextBox txtFields 
         DataField       =   "TOT_ITEM_BORROWED"
         DataSource      =   "datMember"
         Height          =   330
         Index           =   6
         Left            =   1800
         TabIndex        =   7
         Text            =   "0"
         Top             =   3180
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "ACTIVE_STATUS"
         DataSource      =   "datMember"
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   8400
         TabIndex        =   8
         Top             =   3255
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "COURSE_DEPT"
         DataSource      =   "datMember"
         Height          =   315
         Index           =   8
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   6855
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   3
         Left            =   3840
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   795
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Active Status"
         Height          =   225
         Index           =   7
         Left            =   6960
         TabIndex        =   17
         Top             =   3315
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Item Borrowed"
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3315
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone Number"
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2955
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1875
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1215
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "IC Number"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1515
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Member Code"
         Height          =   225
         Index           =   1
         Left            =   5280
         TabIndex        =   11
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Number"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   5160
         TabIndex        =   10
         Top             =   1155
         Width           =   1335
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
      TabIndex        =   25
      Top             =   6240
      Width           =   6615
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   7920
      TabIndex        =   24
      Top             =   6240
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
      Left            =   1920
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
      i1              =   "LMember.frx":0442
   End
   Begin ComctlLib.ImageList imlLargeIcon 
      Left            =   960
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
      i1              =   "LMember.frx":0939
   End
   Begin ComctlLib.ListView lvwLibrary 
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
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
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      SmallIcons      =   "imlSmallIcon"
      Sorted          =   -1  'True
   End
End
Attribute VB_Name = "frmMember"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMember
'***************************************************************************


Private fShowEdit As Boolean


Private Sub cboFields_Click(Index As Integer)
    cboFields(Index).Tag = "Dirty"
End Sub

Private Sub cboFields_GotFocus(Index As Integer)
    lblIndicator = lblLabels(Index) & " ..."
    With frmMDIMainMenu.tbrLibrary
    Select Case Index
        Case 1
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmMemberType
        Case Else
            .Buttons("Properties").Enabled = False
    End Select
    End With
End Sub

Private Sub chkFields_Click(Index As Integer)
    chkFields(Index).Tag = "Dirty"
End Sub

Private Sub cmd3Close_Click()
    Unload Me
End Sub


Public Sub cmd3DBCtrl_Click(Index As Integer)
    If gfSearchMode Then
        Select Case Index
            Case 0
                FieldControlNew Me
                txtFields(6) = 0
                chkFields(7) = vbChecked
            Case 2
                BuildSearchSQL Me
        End Select
    ElseIf Not gfSearchMode Then
        Select Case Index
            Case 0
                txtFields(6) = 0
                chkFields(7) = vbChecked
        End Select
        CtrCommand3D frmMember, Index
    End If
End Sub


Public Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    gtTableActive = "PATRON"
    gtTableIndex = "PATRON_NUM"
    gtMainLvw = "RecPatron"
    
    dbRetrieve cboFields(8), "PATRON", "COURSE_DEPT"
    dbRetrieve cboFields(1), "MEMBER-TYPE", "MEMBER_CODE"
    
    ShowFirst frmMember
    fillLvw Me, 6
    FieldControlDisabled Me
'    WindowState = 2
    CtrEnableLvwViewMain
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gtTableIndex = ""
    CtrEnableLvwViewMain
    gtMainLvw = "MainPatron"
    Init_Report
End Sub


Private Sub lvwLibrary_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  lvwLibrary.SortKey = ColumnHeader.Index - 1
  lvwLibrary.SortOrder = Abs(CInt(gfSort)) ' Ascending
  lvwLibrary.Sorted = 1
End Sub


Private Sub lvwLibrary_ItemClick(ByVal Item As ListItem)
  If getLibSet("SelectRow") = "1" Then
    cSelectRowLvw lvwLibrary
  End If
    gtCurrentIndex = CStr(Item)
    fillCurrentDetail Me, CStr(Item)
End Sub


Private Sub lvwLibrary_MouseDown(Button As Integer, Shift As Integer, _
    X As Single, y As Single)
    If lvwLibrary.ListItems(lvwLibrary.SelectedItem.Index).Selected _
        = True Then
            If (Button = vbRightButton) Then
                PopupMenu frmMDIMainMenu.Edit
            End If
    Else
        If (Button = vbRightButton) Then
            PopupMenu frmMDIMainMenu.View
        End If
    End If
End Sub


Private Sub txtFields_Change(Index As Integer)
    txtFields(Index).Tag = "Dirty"
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    lblIndicator = lblLabels(Index) & " ..."
End Sub




'Private Sub tvwLibrary_MouseMove(Button _
'    As Integer, Shift As Integer, x As _
'    Single, y As Single)
'    Dim AnyNode As Node''
'    ' HitTest returns a node object
'    ' under the cursor
'    Set AnyNode = TreeView1.HitTest(x, _
'        y)
'    If Not AnyNode Is Nothing Then
'        Set TreeView1.DropHighlight = _
'            AnyNode
'        TreeView1.DropHighlight. _
'            Selected = True
'End If
'End Sub

Private Sub lvwLibrary_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'    Dim AnyItem As ListItem

    'HitTest returns a node object under
    'the cursor
'    Set AnyItem = lvwLibrary.HitTest(X,         y)

'    If Not AnyItem Is Nothing Then
'        Set lvwLibrary.DropHighlight =            AnyItem
'        lvwLibrary.DropHighlight.Selected = True
'    End If
End Sub

