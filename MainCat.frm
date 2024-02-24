VERSION 4.00
Begin VB.Form frmMainCataloging 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   975
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7275
   Icon            =   "MainCat.frx":0000
   Left            =   -15
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Top             =   630
   Width           =   7980
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   11625
      TabIndex        =   3
      Top             =   360
      Width           =   11655
      Begin VB.Image imgView 
         Height          =   165
         Left            =   120
         Picture         =   "MainCat.frx":000C
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Frame fraContents 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   11655
      Begin Threed.SSCommand cmd3Close 
         Height          =   300
         Left            =   10080
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
      Begin Threed.SSCommand cmd3GoTo 
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Go to..."
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
   Begin ComctlLib.ImageList imlLargeIcon 
      Left            =   4200
      Top             =   6120
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      NumImages       =   5
      i1              =   "MainCat.frx":046E
      i2              =   "MainCat.frx":0965
      i3              =   "MainCat.frx":0E5C
      i4              =   "MainCat.frx":1353
      i5              =   "MainCat.frx":184A
   End
   Begin ComctlLib.ListView lvwMain 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   8493
      _StockProps     =   205
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Arrange         =   2
      Icons           =   "imlLargeIcon"
      LabelEdit       =   1
      SmallIcons      =   "imlSmallIcon"
   End
   Begin ComctlLib.ImageList imlSmallIcon 
      Left            =   5160
      Top             =   6120
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      NumImages       =   5
      i1              =   "MainCat.frx":1D41
      i2              =   "MainCat.frx":2238
      i3              =   "MainCat.frx":272F
      i4              =   "MainCat.frx":2C26
      i5              =   "MainCat.frx":311D
   End
End
Attribute VB_Name = "frmMainCataloging"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMainCataloging
'***************************************************************************


Private Sub cmd3Close_Click()
    Unload Me
End Sub


Private Sub cmd3GoTo_Click()
    Select Case tSelectedItem
        Case "Item"
            frmMaterial.Show
            DoEvents
        Case "Item Type"
            frmMaterialType.Show
        Case "Placement"
            frmPlacement.Show
        Case "Subject"
            frmSubject.Show
        Case "Language"
            frmLanguage.Show
    End Select
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = 7320
    
    gtMainLvw = "MainCataloging"
    
    Dim itmX As ListItem
    
    Set itmX = lvwMain.ListItems.Add(, "Item", "Item", 1)
        ' Set an icon from ImageList1
        itmX.Icon = 1
        ' Set an icon from ImageList2
        itmX.SmallIcon = 1
    Set itmX = lvwMain.ListItems.Add(, "Item Type", "Item Type", 2)
        ' Set an icon from ImageList1
        itmX.Icon = 2
        ' Set an icon from ImageList2
        itmX.SmallIcon = 2
    Set itmX = lvwMain.ListItems.Add(, "Placement", "Placement", 3)
        ' Set an icon from ImageList1
        itmX.Icon = 3
        ' Set an icon from ImageList2
        itmX.SmallIcon = 3
    Set itmX = lvwMain.ListItems.Add(, "Subject", "Subject", 4)
        ' Set an icon from ImageList1
        itmX.Icon = 4
        ' Set an icon from ImageList2
        itmX.SmallIcon = 4
    Set itmX = lvwMain.ListItems.Add(, "Language", "Language", 5)
        ' Set an icon from ImageList1
        itmX.Icon = 5
        ' Set an icon from ImageList2
        itmX.SmallIcon = 5
    
    CtrEnableLvwViewMain

End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' Disable ChangeView Toolbar
    CtrDisableLvwView
    
    ' Reset variable to detect current Main* Form
    gtMainLvw = ""
End Sub


Private Sub imgView_Click()
    With frmMDIMainMenu
        .Separator8.Visible = False
        .mnuSort.Visible = False
        PopupMenu .View, vbPopupMenuLeftAlign, lvwMain.Left, _
            lvwMain.Top
        .mnuSort.Visible = True
        .Separator8.Visible = True
    End With
End Sub


Private Sub lvwMain_DblClick()
    cmd3GoTo_Click
End Sub


Private Sub lvwMain_ItemClick(ByVal Item As ListItem)
    tSelectedItem = CStr(Item)
End Sub



