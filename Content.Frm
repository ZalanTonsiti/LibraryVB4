VERSION 4.00
Begin VB.Form frmContent 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   510
   ClientTop       =   1140
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7320
   Left            =   450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Top             =   795
   Width           =   9780
   Begin VB.PictureBox picLibrary 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   4905
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Image Image1 
         Height          =   315
         Left            =   120
         Picture         =   "Content.frx":0000
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Chart"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6390
      Left            =   3480
      Top             =   240
      Width           =   4950
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   5
      Left            =   8520
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   4
      Left            =   8520
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Utility"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   3
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Patron"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
      MouseIcon       =   "Content.frx":066E
      Picture         =   "Content.frx":068A
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Acquisition"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Cataloging"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmd3Main 
      Height          =   2055
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   3625
      _StockProps     =   78
      Caption         =   "Circulation"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin GraphLib.Graph grpLibrary 
      Height          =   5895
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   10398
      _StockProps     =   96
      BorderStyle     =   1
      DrawMode        =   1
      GraphType       =   4
      RandomData      =   0
      TickEvery       =   6
      YAxisMax        =   400
      YAxisTicks      =   100
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontFamily[0]   =   1
      FontFamily[1]   =   1
      FontFamily[2]   =   1
      FontFamily[3]   =   1
      FontSize        =   4
      FontSize[0]     =   150
      FontSize[1]     =   100
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      FontStyle[0]    =   2
      FontStyle[1]    =   2
      GraphData       =   0
      GraphData[]     =   5
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   5
   End
End
Attribute VB_Name = "frmContent"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmContent
'***************************************************************************


Private Sub cmd3Main_Click(Index As Integer)
  Select Case Index
    Case 3
        frmMainPatron.Show
    Case 1
        frmMainCataloging.Show
    Case 0
        frmMainCirculation.Show
    Case 2
        frmMainAcquisition.Show
    Case 4
        frmMainUtility.Show
    Case 5
        frmGenReport.Show
  End Select
End Sub


Private Sub Form_GotFocus()
    CtrDisableLvwView
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = frmMDIMainMenu.Height - 1380
    'frmMDIMainMenu.Height - (1380 / Screen.TwipsPerPixelY) '7320
    CtrDisableLvwViewMain
End Sub


Private Sub Form_Resize()
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = frmMDIMainMenu.Height - 1380

End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, _
    X As Single, y As Single)
    If Button = vbLeftButton Then
        PopupMenu frmMDIMainMenu.Graph, , grplibrary.Left, grplibrary.Top
    End If


End Sub


