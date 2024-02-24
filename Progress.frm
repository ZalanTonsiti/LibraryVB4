VERSION 4.00
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   3270
   ClientTop       =   2085
   ClientWidth     =   6690
   Height          =   2280
   Icon            =   "Progress.frx":0000
   Left            =   3210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6690
   Top             =   1740
   Width           =   6810
   Begin VB.TextBox txtInfo 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   6255
   End
   Begin Crystal.CrystalReport repLoanDump 
      Left            =   3000
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ReportFileName  =   ""
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
   Begin Threed.SSCommand cmd3Progress 
      Height          =   300
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "B&rowse"
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
   Begin VB.Label lblInfo 
      Caption         =   "lblInfo"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin Threed.SSCommand cmd3Progress 
      Height          =   300
      Index           =   2
      Left            =   5520
      TabIndex        =   1
      Top             =   1440
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Next >>"
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
   Begin ComctlLib.ProgressBar prgLibrary 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmd3Progress_Click(Index As Integer)
    Select Case Index
    Case 1
        
    Case 2
        Select Case giAnnUpdate
            Case 0
            Case 1
                giAnnUpdate = 2
                dbAnnupdateStepTwo
            Case 2
                ' Print the .. to prn
                repLoanDump.Destination = 0
                repLoanDump.DataFiles(0) = getLibSet("Database")
                repLoanDump.ReportFileName = GetPath & "loan.rpt"
                repLoanDump.SelectionFormula = "{LOAN.CLEARED_FINE} = True"
                repLoanDump.Action = 1
                dbAnnupdateStepThree
            Case 3
                If MsgBox("This stage will delete all celeared fine record permanently. " & _
                    "Click OK to continue", vbExclamation + vbOKCancel, _
                    strsytem & " - WARNING") = vbOK Then
                        SQL = "SELECT * FROM LOAN WHERE CLEARED_FINE = " & True
                        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
                        If Not gLibDyna.EOF Then
                            With gLibDyna
                                While Not .EOF
                                    .Delete
                                    .MoveNext
                                Wend
                             End With
                        End If
                     Else
                    Exit Sub
                 End If
                Unload Me
        End Select
    Case 3
        getFilename
        If txtinfo <> "" Then cmd3Progress(2).Enabled = True
    End Select
End Sub

Private Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
End Sub

Private Sub txtinfo_Change()
    If Len(txtinfo) <> 0 Then gtFileNameAnnUpdt = txtinfo
End Sub
