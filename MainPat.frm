VERSION 4.00
Begin VB.Form frmMainPatron 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   1245
   ClientTop       =   615
   ClientWidth     =   9015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7275
   Left            =   1185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Top             =   270
   Width           =   9135
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.Image imgView 
         Height          =   165
         Left            =   120
         Picture         =   "MainPat.frx":0000
         Top             =   0
         Width           =   195
      End
   End
   Begin ComctlLib.ImageList imlLargeIcon 
      Left            =   1800
      Top             =   6840
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      NumImages       =   2
      i1              =   "MainPat.frx":0462
      i2              =   "MainPat.frx":0959
   End
   Begin ComctlLib.ListView lvwMain 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
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
      Left            =   720
      Top             =   6840
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      NumImages       =   2
      i1              =   "MainPat.frx":0E50
      i2              =   "MainPat.frx":1347
   End
End
Attribute VB_Name = "frmMainPatron"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMainPatron
'***************************************************************************


Private Sub cmd3Close_Click()
    Unload Me
End Sub


Private Sub cmd3GoTo_Click()
    Select Case tSelectedItem
        Case "Member"
            frmMember.Show
        Case "Member Type"
            frmMemberType.Show
    End Select
End Sub


Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = 7320
    gtMainLvw = "MainPatron"
    Dim itmX As ListItem
    Set itmX = lvwMain.ListItems.Add(, "Member", "Member", 1)
        ' Set an icon from ImageList1
        itmX.Icon = 1
        ' Set an icon from ImageList2
        itmX.SmallIcon = 1
    Set itmX = lvwMain.ListItems.Add(, "Member Detail", "Member Type", 2)
        ' Set an icon from ImageList1
        itmX.Icon = 2
        ' Set an icon from ImageList2
        itmX.SmallIcon = 2
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



