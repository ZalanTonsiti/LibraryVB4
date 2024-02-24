VERSION 4.00
Begin VB.Form frmMaterial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item"
   ClientHeight    =   6960
   ClientLeft      =   -15
   ClientTop       =   795
   ClientWidth     =   11895
   Height          =   7365
   Icon            =   "LMaterial.frx":0000
   Left            =   -75
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   Top             =   450
   Width           =   12015
   Begin VB.Frame fraLibrary 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   11655
      Begin VB.TextBox txtFields 
         DataField       =   "ACQUISITION_NUM"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Tag             =   "5"
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "MATERIAL_CODE"
         DataSource      =   "datMaterial"
         Height          =   315
         Index           =   1
         Left            =   5280
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "LANGUAGE_CODE"
         DataSource      =   "datMaterial"
         Height          =   315
         Index           =   2
         Left            =   9120
         TabIndex        =   17
         Top             =   3060
         Width           =   735
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "SHELF_CODE"
         DataSource      =   "datMaterial"
         Height          =   315
         Index           =   3
         Left            =   9120
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "VENDOR_CODE"
         DataSource      =   "datMaterial"
         Height          =   315
         Index           =   0
         Left            =   9120
         TabIndex        =   18
         Top             =   3420
         Width           =   735
      End
      Begin VB.ComboBox cboFields 
         DataField       =   "PUBLISHER_CODE"
         DataSource      =   "datMaterial"
         Height          =   315
         Index           =   4
         Left            =   9120
         TabIndex        =   19
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CLASSIFICATION_NUM"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   7
         Left            =   9120
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "0"
         Top             =   1380
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AUTHOR1"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   8
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1380
         Width           =   6255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AUTHOR2"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   9
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1710
         Width           =   6255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AUTHOR3"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   10
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2040
         Width           =   6255
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ISBN_NUM"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   11
         Left            =   9120
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "0"
         Top             =   1050
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "EDITION"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   12
         Left            =   9120
         TabIndex        =   12
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         DataField       =   "SINOPSIS"
         DataSource      =   "datMaterial"
         Height          =   975
         Index           =   13
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2400
         Width           =   3405
      End
      Begin VB.TextBox txtFields 
         DataField       =   "YEAR_PUBLISHED"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   14
         Left            =   9120
         TabIndex        =   14
         Text            =   "1999"
         Top             =   2730
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PLACE_PUBLISHED"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   15
         Left            =   9120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "ON-LOAN"
         DataSource      =   "datMaterial"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   22
         Top             =   3600
         Width           =   225
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "RESERVE-LIST"
         DataSource      =   "datMaterial"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   3600
         Width           =   225
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "EXISTENCE"
         DataSource      =   "datMaterial"
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "REFERENCE_ONLY"
         DataSource      =   "datMaterial"
         Height          =   210
         Index           =   19
         Left            =   5160
         TabIndex        =   24
         Top             =   3600
         Width           =   225
      End
      Begin VB.TextBox txtFields 
         DataField       =   "TITLE"
         DataSource      =   "datMaterial"
         Height          =   300
         Index           =   20
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   8
         Top             =   1050
         Width           =   6255
      End
      Begin VB.ComboBox cboExistence 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "LMaterial.frx":0442
         Left            =   6120
         List            =   "LMaterial.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3600
         Width           =   1575
      End
      Begin VB.ListBox lstSubject 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   5640
         TabIndex        =   47
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   53
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   3
         Left            =   9960
         TabIndex        =   52
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   2
         Left            =   9960
         TabIndex        =   51
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   1
         Left            =   9960
         TabIndex        =   50
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   0
         Left            =   10080
         TabIndex        =   49
         Top             =   720
         Width           =   1455
      End
      Begin Threed.SSCommand cmd3DBCtrl 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   1
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
         TabIndex        =   2
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
         Left            =   2760
         TabIndex        =   3
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
         Left            =   3840
         TabIndex        =   4
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
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         Height          =   225
         Index           =   20
         Left            =   960
         TabIndex        =   45
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Place Published"
         Height          =   225
         Index           =   15
         Left            =   7800
         TabIndex        =   44
         Top             =   2415
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Year Published"
         Height          =   225
         Index           =   14
         Left            =   7800
         TabIndex        =   43
         Top             =   2745
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sinopsis"
         Height          =   225
         Index           =   13
         Left            =   480
         TabIndex        =   42
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Edition"
         Height          =   225
         Index           =   12
         Left            =   8400
         TabIndex        =   41
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "ISBN Number"
         Height          =   225
         Index           =   11
         Left            =   7800
         TabIndex        =   40
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Author 3"
         Height          =   225
         Index           =   10
         Left            =   600
         TabIndex        =   39
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Author 2"
         Height          =   225
         Index           =   9
         Left            =   600
         TabIndex        =   38
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Author 1"
         Height          =   225
         Index           =   8
         Left            =   600
         TabIndex        =   37
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Class. Number"
         Height          =   225
         Index           =   7
         Left            =   7920
         TabIndex        =   36
         Top             =   1395
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Subject"
         Height          =   225
         Index           =   6
         Left            =   4920
         TabIndex        =   35
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Publisher"
         Height          =   225
         Index           =   5
         Left            =   8280
         TabIndex        =   34
         Top             =   2055
         Width           =   720
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor"
         Height          =   225
         Index           =   4
         Left            =   8400
         TabIndex        =   33
         Top             =   3435
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Placement Shelf"
         Height          =   225
         Index           =   3
         Left            =   7680
         TabIndex        =   32
         Top             =   735
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Language"
         Height          =   225
         Index           =   2
         Left            =   8160
         TabIndex        =   31
         Top             =   3075
         Width           =   840
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Type"
         Height          =   225
         Index           =   1
         Left            =   4080
         TabIndex        =   30
         Top             =   720
         Width           =   1080
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition No"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Booked"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   28
         Top             =   3615
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "On-Loan"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   3615
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Reference Only"
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   3615
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Exist"
         Height          =   255
         Left            =   5640
         TabIndex        =   25
         Top             =   3615
         Width           =   375
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
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   6600
      Width           =   9735
   End
   Begin ComctlLib.ListView lvwLibrary 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
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
      LabelEdit       =   1
      SmallIcons      =   "imlSmallIcon"
      Sorted          =   -1  'True
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   10800
      TabIndex        =   48
      Top             =   6600
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
      i1              =   "LMaterial.frx":0446
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
      i1              =   "LMaterial.frx":093D
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmMaterial
'***************************************************************************


Private Sub cboExistence_Click()
    Select Case cboExistence.ListIndex
        Case 0
            chkFields(2).Value = Checked
        Case 1
            chkFields(2).Value = Unchecked
    End Select

End Sub

Private Sub cboFields_Change(Index As Integer)
    cboFields_Click (Index)
End Sub

Private Sub cboFields_Click(Index As Integer)
    On Local Error GoTo Error_Handler
    Dim SQL
    If cboFields(Index) = "" Then
        Select Case Index
            Case 0
                lblInfo(3) = ""
            Case 1
                lblInfo(4) = ""
            Case 2
                lblInfo(2) = ""
            Case 3
                lblInfo(0) = ""
            Case 4
                lblInfo(1) = ""
            Case 6
                lblInfo(0) = ""
        End Select
        GoTo End_Sub
    End If
    Select Case Index
        Case 0
            SQL = "SELECT VENDOR.VENDOR_NAME FROM VENDOR" & _
            " WHERE VENDOR_CODE= '" & cboFields(0) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("VENDOR_NAME")
                
            lblInfo(3) = gLibSnap.Fields("VENDOR_NAME")
        Case 1
            SQL = "SELECT [MATERIAL-TYPE].MATERIAL_DESC FROM " _
                & "[MATERIAL-TYPE] WHERE MATERIAL_CODE= '" _
                & cboFields(1) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("MATERIAL_DESC")
            lblInfo(4) = gLibSnap.Fields("MATERIAL_DESC")
        Case 2
            SQL = "SELECT LANGUAGE.LANGUAGE_DESC FROM LANGUAGE" & _
            " WHERE LANGUAGE_CODE= '" & cboFields(2) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("LANGUAGE_DESC")
            lblInfo(2) = gLibSnap.Fields("LANGUAGE_DESC")
            
        Case 3
            SQL = "SELECT PLACEMENT.SHELF_DESC FROM PLACEMENT" & _
            " WHERE SHELF_CODE= '" & cboFields(3) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("SHELF_DESC")
            lblInfo(0) = gLibSnap.Fields("SHELF_DESC")
        Case 4
            SQL = "SELECT PUBLISHER.PUBLISHER_NAME FROM PUBLISHER" & _
            " WHERE PUBLISHER_CODE= '" & cboFields(4) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("PUBLISHER_NAME")
            lblInfo(1) = gLibSnap.Fields("PUBLISHER_NAME")
        Case 6
            SQL = "SELECT SUBJECT.SUBJECT_DESC FROM SUBJECT" & _
            " WHERE SUBJECT_CODE= '" & cboFields(6) & "';"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = _
                gLibSnap.Fields("SUBJECT_DESC")
            lblInfo(0) = gLibSnap.Fields("SUBJECT_DESC")
    End Select
    cboFields(Index).Tag = "Dirty"
    GoTo End_Sub
Error_Handler:
    DisplayError ""
    Resume End_Sub
End_Sub:
End Sub


Private Sub cboFields_GotFocus(Index As Integer)
    cboFields_Click (Index)
    With frmMDIMainMenu.tbrLibrary
    Select Case Index
        Case 0
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmVendor
        Case 1
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmMaterialType
        Case 2
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmLanguage
        Case 3
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmPlacement
        Case 4
            .Buttons("Properties").Enabled = True
            Set gCtrFrm = frmPublisher
        Case Else
            .Buttons("Properties").Enabled = False
    End Select
    End With
End Sub


Private Sub cboFields_LostFocus(Index As Integer)
    frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = ""
End Sub


Private Sub chkFields_Click(Index As Integer)
    chkFields(Index).Tag = "Dirty"
Select Case chkFields(2).Value
    Case Unchecked
        cboExistence.ListIndex = 1
    Case Checked
            cboExistence.ListIndex = 0
End Select
End Sub

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
        CtrCommand3D frmMaterial, Index
    End If

    
    ' User Interface control (for this form only)
    Select Case Index
        Case 0
            chkFields(2).Value = Checked
            lstSubject.Enabled = True
            lstSubject.Clear
            cboExistence.Enabled = True
        Case 1
            lstSubject.Enabled = True
            cboExistence.Enabled = True
        
        Case 2, 3
            If Not gfSearchMode Then
                lstSubject.Enabled = False
                cboExistence.Enabled = False
            Else
                lstSubject.Enabled = True
                cboExistence.Enabled = True
            End If
    End Select
End Sub


Public Sub Form_Load()
'    WindowState = 2
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = 7320

    dbRetrieve cboFields(1), "MATERIAL-TYPE", "MATERIAL_CODE"
    dbRetrieve cboFields(4), "PUBLISHER", "PUBLISHER_CODE"
    dbRetrieve cboFields(2), "LANGUAGE", "LANGUAGE_CODE"
    dbRetrieve cboFields(3), "PLACEMENT", "SHELF_CODE"
    dbRetrieve cboFields(0), "VENDOR", "VENDOR_CODE"
    cboExistence.AddItem "Available"
    cboExistence.AddItem "Missing"
    gtTableActive = "MATERIAL"
    gtTableIndex = "ACQUISITION_NUM"
    gtMainLvw = "RecMaterial"
    fillLvw Me, 5
    FieldControlDisabled Me
    fillCurrentDetail Me, 75
    fillLstSubject 75
    gtCurrentIndex = 75
    MsgBar "", False
    Form_Resize
    CtrEnableLvwViewMain
End Sub


Public Sub Form_Resize()
    If WindowState <> 1 Then
        Me.Left = 0
        Me.Top = 0
        Me.Width = 11930
        Me.Height = 7320
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gtTableIndex = ""
    CtrEnableLvwViewMain
    gtMainLvw = "MainCataloging"
End Sub


Private Sub lblInfo_Change(Index As Integer)
    If Len(lblInfo(Index)) > 15 Then
        lblInfo(Index) = Left(lblInfo(Index), 12) & " ..."
    End If
End Sub

Private Sub lstSubject_DblClick()
    If gfSearchMode Then
        frmSearchSubject.Show
    Else
        frmSubjectMat.Show
    End If
End Sub


Private Sub lvwLibrary_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  lvwLibrary.SortKey = ColumnHeader.Index - 1
  lvwLibrary.SortOrder = Abs(CInt(gfSort))      ' Ascending
  lvwLibrary.Sorted = 1
End Sub


Private Sub lvwLibrary_ItemClick(ByVal Item As ListItem)
    gtAcquisition_Num = CStr(Item)
    gtCurrentIndex = CStr(Item)
    fillCurrentDetail Me, CSng(Item)
    fillLstSubject CSng(Item)
  If getLibSet("SelectRow") = "1" Then
    cSelectRowLvw lvwLibrary
  End If
    frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = ""
End Sub


'===========================================================================
'   DESC    :   Fill List Control with SUBJECTs
'===========================================================================
Private Sub fillLstSubject(sID As Single)
    lstSubject.Clear
    Dim tSubject As String
    SQL = "SELECT MATERIAL.ACQUISITION_NUM, [SUBJECT-MATERIAL].SUBJECT_CODE," _
        & " SUBJECT.SUBJECT_DESC FROM MATERIAL INNER JOIN (SUBJECT INNER " _
        & "JOIN [SUBJECT-MATERIAL] ON SUBJECT.SUBJECT_CODE = " _
        & "[SUBJECT-MATERIAL].SUBJECT_CODE) ON MATERIAL.ACQUISITION_NUM " _
        & "= [SUBJECT-MATERIAL].ACQUISITION_NUM WHERE " _
        & "MATERIAL.ACQUISITION_NUM = '" & sID & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    If Not gLibSnap.EOF Then
    While Not gLibSnap.EOF
        tSubject = CStr(gLibSnap.Fields(1)) & ", " & _
            CStr(gLibSnap.Fields(2))
        lstSubject.AddItem tSubject
        gLibSnap.MoveNext
    Wend
    Else
        lstSubject.AddItem ""
    End If
End Sub


Private Sub txtFields_Change(Index As Integer)
    txtFields(Index).Tag = "Dirty"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        Select Case Index       ' Optimize - Use frm.Controls(index+1)
            Case 0
                cboFields(1).SetFocus
             Case 20
                txtFields(8).SetFocus
             Case 8
                txtFields(Index + 1).SetFocus
             Case 9
                txtFields(Index + 1).SetFocus
            Case 10
                txtFields(13).SetFocus
            Case 13
                lstSubject.SetFocus
            Case 11
                txtFields(7).SetFocus
            Case 7
                txtFields(12).SetFocus
            Case 12
                cboFields(4).SetFocus
            Case 15
                txtFields(14).SetFocus
            Case 14
                cboFields(2).SetFocus
          End Select
    KeyAscii = 0
    End If



    Select Case Index
        Case 0, 14, 12
            If Not (chkIsNum(Chr$(KeyAscii))) Then
                MsgBox "This field require numeric value", vbOKOnly + vbInformation, strSystem
                txtFields(Index).Text = ""
            End If
    End Select




End Sub




Private Sub cboFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0
            cmd3DBCtrl(2).SetFocus
            Case 1
                txtFields(20).SetFocus
            Case 2
                cboFields(0).SetFocus
            Case 3
                txtFields(11).SetFocus
            Case 4
                txtFields(15).SetFocus
    
        End Select
    KeyAscii = 0
    End If
    
End Sub

