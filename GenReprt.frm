VERSION 4.00
Begin VB.Form frmGenReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Report"
   ClientHeight    =   7755
   ClientLeft      =   90
   ClientTop       =   1050
   ClientWidth     =   11880
   Height          =   8160
   Icon            =   "GenReprt.frx":0000
   Left            =   30
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Top             =   705
   Width           =   12000
   Begin VB.ListBox lstReportList 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
   End
   Begin Threed.SSCommand cmd3GenReport 
      Height          =   300
      Index           =   1
      Left            =   10680
      TabIndex        =   3
      Top             =   6000
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
   Begin Threed.SSCommand cmd3GenReport 
      Height          =   300
      Index           =   0
      Left            =   9480
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Generate..."
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
   Begin Crystal.CrystalReport repLibrary 
      Left            =   2160
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ReportFileName  =   "d:\lib\member cards.rpt"
      Destination     =   0
      WindowLeft      =   100
      WindowTop       =   100
      WindowWidth     =   490
      WindowHeight    =   300
      WindowTitle     =   ""
      WindowBorderStyle=   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   1
      PrintFileName   =   ""
      PrintFileType   =   0
      SelectionFormula=   ""
      GroupSelectionFormula=   ""
      Connect         =   ""
      UserName        =   ""
      ReportSource    =   0
      BoundReportHeading=   ""
      BoundReportFooter=   0   'False
   End
   Begin VB.Label lblInfo 
      Caption         =   "Please choose the report to generate: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGenReport"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmGenReport
'***************************************************************************


Public Sub cmd3GenReport_Click(Index As Integer)
    On Error GoTo ErrGenerate
    Dim pSelFormula As String
    Select Case lstReportList.ListIndex + 1
        Case 10
            gtRepSearchItem = "{LOAN.DATE_CLEARED_FINE}"
            gtTableActive = "LOAN"
        Case 11
            gtRepSearchItem = "{SUPPLY.DATE_OF_SUPPLY}"
            gtTableActive = "SUPPLY"
    End Select
    
    Select Case Index
        Case 0
            Select Case gReportFile(lstReportList.ListIndex + 1).Parameter
                Case 1
                Case 2  ' Show select PATRON_NUM custom CommonDialog
                    frmGetPatronNum.Show 1
                    Init_Report     ' ReInit
                Case 3
                If CheckReportDate Then
                    InitDayMonth
                    frmGetDate.Show 1
                    Init_Report
                Else: Exit Sub
                End If
            End Select
            
            pSelFormula = gReportFile(lstReportList.ListIndex + 1).SelFormula
            replibrary.Destination = 0
            replibrary.DataFiles(0) = gDbName$
            replibrary.ReportFileName = _
            GetPath & Trim$(gReportFile(lstReportList.ListIndex _
                + 1).Name) & ".rpt"
            
            If Len(pSelFormula) > 0 Then
                replibrary.SelectionFormula = pSelFormula
            Else
                replibrary.SelectionFormula = ""
            End If
            
            Debug.Print "replibrary.SelectionFormula >> "; replibrary.SelectionFormula
            replibrary.Action = 1
        Case 1
            Unload Me
    End Select
    Exit Sub
ErrGenerate:
    Select Case Err.Number
    Case 20534
        MsgBox Error$ & " The database DLL is corrupt."
    Case Else
        MsgBox Error$
    End Select
    Exit Sub
End Sub


Private Sub Form_Load()
    gtMainLvw = "GenReport"
    Me.Left = 0
    Me.Top = 0
    Me.Width = 11930
    Me.Height = 7320
    
    Dim Index
    With lstReportList
        For Index = LBound(gReportFile) To UBound(gReportFile)
            .AddItem gReportFile(Index).Name
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMDIMainMenu.sbrLibrary.Panels.Item(1).Text = ""
   frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = ""
   gtMainLvw = ""
   gtTableActive = ""
End Sub

Private Sub lstReportList_Click()
   frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = CStr(GetPath & lstReportList & ".rpt")
    
    If Dir(CStr(GetPath & lstReportList & ".rpt")) <> "" Then
       frmMDIMainMenu.sbrLibrary.Panels.Item(1).Text = "Exist"
    Else
       frmMDIMainMenu.sbrLibrary.Panels.Item(1).Text = "Not Exist"
    End If

End Sub

Private Sub lstReportList_DblClick()
    cmd3GenReport_Click (0)
End Sub
