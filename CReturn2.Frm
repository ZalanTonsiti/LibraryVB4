VERSION 4.00
Begin VB.Form frmReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return"
   ClientHeight    =   6390
   ClientLeft      =   1110
   ClientTop       =   1620
   ClientWidth     =   9720
   Height          =   6795
   Icon            =   "CReturn2.frx":0000
   Left            =   1050
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   0
   ScaleWidth      =   0
   Top             =   1275
   Width           =   9840
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   9495
      Begin VB.TextBox txtAcqNum 
         Height          =   285
         Left            =   6360
         TabIndex        =   14
         Top             =   960
         Width           =   3015
      End
      Begin VB.ListBox lstAcqNum 
         Height          =   1035
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtPatronNum 
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.ListBox lstPatronNum 
         Height          =   1035
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblMaterialTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6360
         TabIndex        =   17
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblInfoAcqName 
         Alignment       =   1  'Right Justify
         Caption         =   "Material Title"
         Height          =   255
         Left            =   5280
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblInoAcqNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition Number"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   240
         Left            =   840
         TabIndex        =   9
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Number"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin Threed.SSCommand cmd3Return 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Return"
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
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Fine"
         Height          =   255
         Left            =   7200
         TabIndex        =   6
         Top             =   255
         Width           =   735
      End
      Begin Threed.SSCommand cmd3ClearFine 
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Clear Fine"
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
      Begin VB.Label lblPatronName 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8040
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin ComctlLib.ImageList imlTvw 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      NumImages       =   2
      i1              =   "CReturn2.frx":0442
      i2              =   "CReturn2.frx":0939
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   8640
      TabIndex        =   10
      Top             =   6000
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
   Begin ComctlLib.TreeView tvwLibrary 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   4471
      _StockProps     =   196
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
      ImageList       =   "imlTvw"
      Indentation     =   12
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "\"
      Style           =   7
   End
   Begin ComctlLib.ListView lvwReturn 
      Height          =   2535
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   4471
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
      Icons           =   "imlTvw"
      SmallIcons      =   ""
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


'***************************************************************************
' frmReturn
'***************************************************************************

Private tCurrentAcquisition As String    ' Current MATERIAL selected


Private Sub fillLoan()
    On Error GoTo Error_Handler:
    If Len(txtPatronNum) Then
        SQL = "SELECT LOAN.PATRON_NUM, LOAN.ACQUISITION_NUM, PATRON.NAME, " & _
                "MATERIAL.TITLE, LOAN.DATE_BORROW, LOAN.TOT_FINE, LOAN.DATE_RETURN, " & _
                "LOAN.CLEARED_FINE FROM PATRON INNER JOIN (MATERIAL INNER JOIN LOAN " & _
                "ON MATERIAL.ACQUISITION_NUM = LOAN.ACQUISITION_NUM) ON " & _
                "PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE LOAN.PATRON_NUM = '" _
                & txtPatronNum.Text & "';"
    
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    If gLibSnap.EOF Then
        lblPatronName = ""
    Else
        lblPatronName = gLibSnap.Fields("NAME")
        
        If gLibSnap.EOF Then
            clearLvw
        Else
            fillLvwReturn txtPatronNum
        End If
    
    End If
  
  End If

Exit Sub

Error_Handler:
    DisplayError ""

End Sub


Private Sub cmd3ClearFine_Click()
    If tCurrentAcquisition = "" Then Exit Sub
    ' Set Cleared Fine for a material
    SQL = "SELECT CLEARED_FINE, DATE_CLEARED_FINE FROM LOAN WHERE ACQUISITION_NUM = '" & _
          tCurrentAcquisition & "' ORDER BY DATE_BORROW DESC"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    With gLibDyna
    If Not .EOF Then
        .Edit
            .Fields("CLEARED_FINE") = True
            .Fields("DATE_CLEARED_FINE") = Date
        .Update
    End If
    End With
    
    fillLvwReturn txtPatronNum.Text
    
    ' Update ACTIVE STATUS  of PATRON if Total Borrowed < Max Borrow Allowed and Total Fine=0
        SQL = "SELECT PATRON.TOT_ITEM_BORROWED, [MEMBER-TYPE].MAX_BORROWING, PATRON.ACTIVE_STATUS " & _
              "FROM [MEMBER-TYPE] INNER JOIN PATRON ON [MEMBER-TYPE].MEMBER_CODE = PATRON.MEMBER_CODE " & _
              "WHERE PATRON_NUM = '" & txtPatronNum & "'"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        With gLibDyna
        If .Fields("MAX_BORROWING") > .Fields("TOT_ITEM_BORROWED") And (CSng(lblTotal) = 0) Then
            .Edit
                .Fields("ACTIVE_STATUS") = True
            .Update
        End If
        End With
End Sub


Private Sub cmd3Close_Click()
    Unload Me
End Sub

Private Sub cmd3Return_Click()
    On Local Error GoTo Error_Handler
    If txtPatronNum.Text = "" Then Exit Sub
    Dim cTotalFine As Currency
    Dim xDateBorrow As Date
    Dim iBorrowDuration As Integer
    Dim sFineRate As Single
    Dim iClaimDuration As Integer
    
    SQL = "SELECT DATE_BORROW FROM LOAN WHERE ACQUISITION_NUM = '" & _
        tCurrentAcquisition & "' ORDER BY DATE_BORROW DESC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    If Date < gLibSnap.Fields(0) Then
        MsgBox "Date borrowed is latter than Date Return.", vbCritical, strSystem
        Exit Sub
    End If
    
    ' Update DATE_RETURN
    SQL = "SELECT DATE_RETURN FROM LOAN WHERE ACQUISITION_NUM = '" & _
        tCurrentAcquisition & "' ORDER BY DATE_BORROW DESC"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    With gLibDyna
    .Edit
        If IsNull(.Fields("DATE_RETURN")) Then
            .Fields("DATE_RETURN") = Date
        Else
            MsgBox "The material has been returned.", vbInformation
            .CancelUpdate
            Exit Sub
        End If
    .Update
    End With
    
    fillLvwReturn txtPatronNum.Text
    
    ' Update LOAN_TABLE
        ' Calculate TOT_FINE
        
    SQL = "SELECT PATRON.PATRON_NUM, [MATERIAL-TYPE].DURATION_OF_BORROW, " & _
        "LOAN.ACQUISITION_NUM, LOAN.DATE_BORROW, LOAN.DATE_RETURN, " & _
        "[MEMBER-TYPE].FINE_RATE_PERDAY FROM ([MEMBER-TYPE] INNER JOIN " & _
        "PATRON ON [MEMBER-TYPE].MEMBER_CODE = PATRON.MEMBER_CODE) INNER JOIN " & _
        "(([MATERIAL-TYPE] INNER JOIN MATERIAL ON [MATERIAL-TYPE].MATERIAL_CODE " & _
        "= MATERIAL.MATERIAL_CODE) INNER JOIN LOAN ON MATERIAL.ACQUISITION_NUM = " & _
        "LOAN.ACQUISITION_NUM) ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE " & _
        "LOAN.ACQUISITION_NUM = '" & tCurrentAcquisition & "' AND LOAN.PATRON_NUM = '" & _
        txtPatronNum.Text & "' AND LOAN.DATE_RETURN = #" & Date & "# ORDER BY DATE_RETURN DESC;"
        
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        
        iBorrowDuration = gLibSnap.Fields("DURATION_OF_BORROW")
        xDateBorrow = gLibSnap.Fields("DATE_BORROW")
        sFineRate = gLibSnap.Fields("FINE_RATE_PERDAY")
        
        SQL = "SELECT LOAN.TOT_FINE, LOAN.CLEARED_FINE FROM LOAN WHERE ACQUISITION_NUM = '" & _
              tCurrentAcquisition & "' AND PATRON_NUM = '" & txtPatronNum.Text & _
              "' AND DATE_RETURN = #" & Date & "# ORDER BY DATE_RETURN DESC"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        
        ' Becareful with date precedence
        If xDateBorrow + iBorrowDuration < Date Then
            cTotalFine = (Date - (xDateBorrow + iBorrowDuration)) * sFineRate
            With gLibDyna
                .Edit
                    .Fields("TOT_FINE") = cTotalFine
                .Update
            End With
            ' Update Active_Status of Patron
            SQL = "SELECT ACTIVE_STATUS FROM PATRON WHERE PATRON_NUM = '" & txtPatronNum.Text & "'"
            Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            With gLibDyna
                .Edit
                    .Fields(0) = False
                .Update
            End With
                    
            fillLvwReturn txtPatronNum.Text
        Else                ' The book returned with borrowing duration allowed
            With gLibDyna
                .Edit
                    .Fields("CLEARED_FINE") = True
                .Update
            End With
            fillLvwReturn txtPatronNum.Text
        End If
    
    ' Update TOTAL ITEM BORROWED by decrement 1
    SQL = "SELECT TOT_ITEM_BORROWED FROM PATRON WHERE PATRON_NUM = '" & txtPatronNum & "'"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    With gLibDyna
        .Edit
            .Fields("TOT_ITEM_BORROWED") = CInt(.Fields("TOT_ITEM_BORROWED")) - 1
        .Update
    End With
    
    ' Update ON-LOAN status of MATERIAL
    SQL = "SELECT [ON-LOAN] FROM MATERIAL WHERE ACQUISITION_NUM = '" & tCurrentAcquisition & "'"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    With gLibDyna
        .Edit
            .Fields("ON-LOAN") = False
        .Update
    End With
    
    ' Check MATERIAL on RESERVE TABLE
    SQL = "SELECT RESERVE.ACQUISITION_NUM FROM RESERVE WHERE ACQUISITION_NUM = '" & tCurrentAcquisition & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    ' Update DATE_OBTAIN
    ' Calculate EXPIRY_BOOKING_DATE
    
    If Not gLibSnap.EOF Then
        iClaimDuration% = CInt(getLibSet("ClaimDay"))
        SQL = "SELECT RESERVE.RECEIVE_DATE, RESERVE.EXPIRY_BOOKING_DATE FROM RESERVE WHERE ACQUISITION_NUM = '" & tCurrentAcquisition & "'"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        With gLibDyna
            .Edit
                .Fields("RECEIVE_DATE") = Date
                .Fields("EXPIRY_BOOKING_DATE") = Date + iClaimDuration%
            .Update
        End With
    End If
    
    ' Generate MEMO to PATRON
    SQL = "SELECT RESERVE.PATRON_NUM, PATRON.NAME, RESERVE.ACQUISITION_NUM, " & _
        "MATERIAL.TITLE, RESERVE.RECEIVE_DATE, RESERVE.EXPIRY_BOOKING_DATE " & _
        "FROM MATERIAL INNER JOIN (PATRON INNER JOIN RESERVE ON PATRON.PATRON_NUM " & _
        "= RESERVE.PATRON_NUM) ON MATERIAL.ACQUISITION_NUM = RESERVE.ACQUISITION_NUM " & _
        "WHERE RESERVE.ACQUISITION_NUM = '" & tCurrentAcquisition & "' ORDER BY " & _
        "EXPIRY_BOOKING_DATE DESC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)

    If Not gLibSnap.EOF Then
      With frmGenReport.replibrary
        .Destination = 0
        .DataFiles(0) = gDbName$
        .ReportFileName = GetPath & "Reserve.rpt"
        .SelectionFormula = "{Reserve Detail.ACQUISITION_NUM} = '" & tCurrentAcquisition & "'"
        .Action = 1
      End With
    End If
    GoTo End_Sub
Error_Handler:
    DisplayError ""
    Resume End_Sub
End_Sub:
End Sub


Private Sub Form_GotFocus()
    Form_Load
End Sub


Private Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    fillTvw
    fillLvwPatronBorrow
    dbRetrieve lstPatronNum, "PATRON", "PATRON_NUM", 1
    frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = "Double click Item to select"
    dbRetrieveMaterialLoan
End Sub


Private Sub lblMaterialTitle_Change()
    If lblMaterialTitle <> "" Then
        tCurrentAcquisition = txtAcqNum
        SQL = "SELECT PATRON_NUM FROM LOAN WHERE ACQUISITION_NUM = '" & txtAcqNum & "'"
        '            " ORDER BY DATE_BORROW DESC"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        
        If Not gLibSnap.EOF Then
            srchLstItem Me, lstPatronNum, CStr(gLibSnap.Fields(0))
            lstPatronNum_DblClick
        End If
    End If
End Sub


Private Sub lblTotal_Change()
    If CSng(lblTotal) > 0 Then
        cmd3ClearFine.Enabled = True
    Else
        cmd3ClearFine.Enabled = False
    End If
End Sub


Private Sub lstAcqNum_DblClick()
    txtAcqNum = lstAcqNum
End Sub


Private Sub lstPatronNum_DblClick()
    fillLvwReturn lstPatronNum
    DoEvents
    txtPatronNum = lstPatronNum
End Sub


Private Sub lvwReturn_ItemClick(ByVal Item As ListItem)
    tCurrentAcquisition = Item
End Sub


' Procedure to fill the ListView Control
' Made Public so that other form could execute it
Public Sub fillLvwReturn(ptPatronNum As String)
    Dim pLibStatus As Recordset
    Dim colColumn As ColumnHeader
    Dim colItem As ListItem
    Dim iIndex As Integer
    Dim sTotalFine As Single
    Dim sAcqNum As Single
    sTotalFine = 0
    
    clearLvw
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "ACQUISITION NO.", "ACQUISITION NO.", lvwReturn.Width / 20)
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "MATERIAL TITLE", "MATERIAL TITLE", lvwReturn.Width / 4)
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "DATE BORROW", "DATE BORROW", lvwReturn.Width / 10)
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "DATE RETURN", "DATE RETURN", lvwReturn.Width / 10)
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "FINE RM", "FINE RM", lvwReturn.Width / 14)
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "CLEARED FINE", "CLEARED FINE", lvwReturn.Width / 14)
    DoEvents
    SQL = "SELECT LOAN.PATRON_NUM, LOAN.ACQUISITION_NUM, PATRON.NAME, " & _
          "MATERIAL.TITLE, LOAN.DATE_BORROW, LOAN.TOT_FINE, LOAN.DATE_RETURN, " & _
          "LOAN.CLEARED_FINE FROM PATRON INNER JOIN (MATERIAL INNER JOIN LOAN " & _
          "ON MATERIAL.ACQUISITION_NUM = LOAN.ACQUISITION_NUM) ON " & _
          "PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE LOAN.PATRON_NUM = '" _
          & ptPatronNum & "' ORDER BY DATE_RETURN ASC;"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    While Not gLibSnap.EOF
      If Not gLibSnap.Fields("CLEARED_FINE") Then   ' Filter unReturned Material Only
        Set colItem = lvwReturn.ListItems.Add(, , CStr(gLibSnap.Fields("ACQUISITION_NUM")))
        sAcqNum = gLibSnap.Fields("ACQUISITION_NUM")
        If Not IsNull(gLibSnap.Fields("TITLE")) Then
            colItem.SubItems(1) = CStr(gLibSnap.Fields("TITLE"))
        End If
                If Not IsNull(gLibSnap.Fields("DATE_BORROW")) Then
                    colItem.SubItems(2) = CStr(gLibSnap.Fields("DATE_BORROW"))
                End If
                If Not IsNull(gLibSnap.Fields("DATE_RETURN")) Then
                    colItem.SubItems(3) = CStr(gLibSnap.Fields("DATE_RETURN"))
'                Else
'                    colItem.SubItems(5) = "ON-LOAN"
                End If
                If Not IsNull(gLibSnap.Fields("TOT_FINE")) Then
                    colItem.SubItems(4) = Format(CStr(gLibSnap.Fields("TOT_FINE")), "###,##0.00")
                End If
                If gLibSnap.Fields("CLEARED_FINE") = False Then ' Optimize ConvBool
                    colItem.SubItems(5) = "No"
                    sTotalFine = sTotalFine + CSng(gLibSnap.Fields("TOT_FINE"))
                Else
                    colItem.SubItems(5) = "Yes"
                End If
      End If
                gLibSnap.MoveNext
    Wend
    lblTotal = Format(sTotalFine, "###,##0.00")
End Sub


' Clear ListView control and set view to viewReport
Private Sub clearLvw()
  lvwReturn.View = lvwReport
  lvwReturn.ColumnHeaders.Clear
  lvwReturn.ListItems.Clear
End Sub


Private Sub fillTvw()
    Dim colNode As Node
    tvwLibrary.Nodes.Clear
    SQL = "SELECT DISTINCT LOAN.PATRON_NUM, PATRON.TOT_ITEM_BORROWED, LOAN.CLEARED_FINE FROM " & _
        "PATRON INNER JOIN LOAN ON PATRON.PATRON_NUM = LOAN.PATRON_NUM " & _
        "WHERE LOAN.CLEARED_FINE = " & False & " ORDER BY LOAN.PATRON_NUM ASC"
    
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    
        Set colNode = tvwLibrary.Nodes.Add(, , "Root", "Patron", 1)
        While Not gLibSnap.EOF
        'Or CInt(gLibSnap.Fields("CLEARED_FINE"))= False
          If CInt(gLibSnap.Fields("TOT_ITEM_BORROWED")) > 0 Then
            Set colNode = tvwLibrary.Nodes.Add("Root", tvwChild, "C" _
                & CStr(gLibSnap.Fields("PATRON_NUM")), CStr(gLibSnap.Fields("PATRON_NUM")), 2)
Debug.Print gLibSnap.Fields("PATRON_NUM")
            colNode.EnsureVisible
          End If
          gLibSnap.MoveNext
        Wend
End Sub


Private Sub tvwLibrary_NodeClick(ByVal Node As Node)
    If Parse(Node.FullPath, 2, "\") <> "" Then
        txtPatronNum = Parse(Node.FullPath, 2, "\")
    Else
        fillLvwPatronBorrow
    End If
End Sub


Private Sub txtAcqNum_Change()
    srchLikeLstItem Me, lstAcqNum, txtAcqNum
    On Error GoTo Error_Handler
    If Len(txtAcqNum) Then
        SQL = "SELECT TITLE FROM MATERIAL WHERE ACQUISITION_NUM = '" _
            & txtAcqNum & "'"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If gLibSnap.EOF Then
            lblMaterialTitle = ""
        Else
            If fSrchLstItem(Me, lstAcqNum, txtAcqNum) Then
                lblMaterialTitle = gLibSnap.Fields("TITLE")
            Else
                lblMaterialTitle = ""
            End If
        End If
    Else
        lblMaterialTitle = ""
    End If
Exit Sub
Error_Handler:
    DisplayError ""

End Sub


Private Sub txtPatronNum_Change()
    srchLikeLstItem Me, lstPatronNum, txtPatronNum
    fillLoan
End Sub


Private Sub fillLvwPatronBorrow()
    
    Dim colColumn As ColumnHeader
    Dim colItem As ListItem
    Dim iIndex As Integer
    clearLvw
    lvwReturn.View = lvwIcon
    
    SQL = "SELECT DISTINCT LOAN.PATRON_NUM, PATRON.TOT_ITEM_BORROWED FROM PATRON INNER JOIN LOAN ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE TOT_ITEM_BORROWED > 0 ORDER BY LOAN.PATRON_NUM ASC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)

    ' Display Not Returned Material or Cleared_Fine = False
    Set colColumn = lvwReturn.ColumnHeaders.Add(, "PATRON NO.", "PATRON NO.", lvwReturn.Width / 3)
    While Not gLibSnap.EOF
        Set colItem = lvwReturn.ListItems.Add(, , CStr(gLibSnap.Fields("PATRON_NUM")), 2)
        gLibSnap.MoveNext
    Wend
    
    AdjustColumnWidth lvwReturn, True

End Sub


' Show On-Loan Material
Private Sub dbRetrieveMaterialLoan()
    lstAcqNum.Clear
    SQL = "SELECT DISTINCT MATERIAL.ACQUISITION_NUM FROM MATERIAL WHERE [ON-LOAN] = " & True
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    While Not gLibSnap.EOF
        If Not IsNull(gLibSnap.Fields(0)) Then
            lstAcqNum.AddItem gLibSnap.Fields(0)
        End If
        gLibSnap.MoveNext
    Wend
    If lstAcqNum.ListCount Then lstAcqNum.ListIndex = 0
End Sub

