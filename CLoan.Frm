VERSION 4.00
Begin VB.Form frmLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan"
   ClientHeight    =   6375
   ClientLeft      =   1785
   ClientTop       =   1290
   ClientWidth     =   9735
   Height          =   6780
   Icon            =   "CLoan.frx":0000
   Left            =   1725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   0
   ScaleWidth      =   0
   Top             =   945
   Width           =   9855
   Begin VB.Frame fraLoan 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9495
      Begin VB.ListBox lstAcqNum 
         Height          =   1620
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ListBox lstPatronNum 
         Height          =   1620
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtPatronNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtAcqNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin Threed.SSCommand cmd3LoanUpdate 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Update"
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
      Begin Threed.SSCommand cmd3AddAnother 
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Add A&nother"
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
      Begin VB.Label lblInfoPatronNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Number"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblInfoPatronName 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Name"
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblInoAcqNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition Number"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblInfoAcqName 
         Alignment       =   1  'Right Justify
         Caption         =   "Material Title"
         Height          =   240
         Left            =   5280
         TabIndex        =   11
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblPatronName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label lblMaterialTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   3240
         Width           =   3015
      End
   End
   Begin ComctlLib.ImageList imlTvw 
      Left            =   5880
      Top             =   6840
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      NumImages       =   2
      i1              =   "CLoan.frx":0442
      i2              =   "CLoan.frx":0939
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
      TabIndex        =   16
      Top             =   6000
      Width           =   6135
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   8520
      TabIndex        =   7
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
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   3625
      _StockProps     =   196
      Appearance      =   1
      ImageList       =   "imlTvw"
      Indentation     =   12
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "\"
      Style           =   7
   End
   Begin ComctlLib.ListView lvwLoan 
      Height          =   2055
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   3625
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
Attribute VB_Name = "frmLoan"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


'***************************************************************************
' frmLoan
'***************************************************************************

Private tPrevious           As String
Private tSelectedItem       As String
Private fUpdateLoanP        As Boolean
Private fUpdateLoanA        As Boolean
Private fLoanEnabled        As Boolean


Private Sub cmd3AddAnother_Click()
    txtPatronNum = tPrevious
    Debug.Print "cmd3AddAnother_Click"
    CheckActiveStatus txtPatronNum
    txtAcqNum.SetFocus
End Sub


Private Sub cmd3Close_Click()
    Unload Me
End Sub

Private Sub cmd3LoanUpdate_Click()
    If (Len(txtPatronNum) = 0) Or (Len(txtAcqNum) = 0) Then Exit Sub
    On Local Error GoTo Error_Handler
    
    CheckActiveStatus txtPatronNum
    
    If fLoanEnabled = False Then Exit Sub
    
    ' Check MATERIAL status
    SQL = "SELECT EXISTENCE, REFERENCE_ONLY, [ON-LOAN] FROM MATERIAL " _
        & "WHERE ACQUISITION_NUM = '" & txtAcqNum & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        With gLibSnap
            fValidBorrow = .Fields("ON-LOAN") Or Not (.Fields("EXISTENCE")) _
                Or .Fields("REFERENCE_ONLY")
        End With
        If fValidBorrow Then
            MsgBox "The material to be borrowed has been borrowed, " _
                & "LOST or RFERENCE ONLY", vbExclamation
            txtAcqNum = ""
            txtAcqNum.SetFocus
            Exit Sub
        End If
        txtPatronNum_Change
        
        ' Update LOAN table
        SQL = "SELECT * FROM LOAN"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            With gLibDyna
                .AddNew
                    .Fields("ACQUISITION_NUM") = txtAcqNum
                    .Fields("PATRON_NUM") = txtPatronNum
                    .Fields("DATE_BORROW") = Date
                .Update
            End With
        
        ' Update ON-LOAN status
        SQL = "SELECT * FROM MATERIAL WHERE ACQUISITION_NUM = '" & txtAcqNum & "'"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        With gLibDyna
            .Edit
                .Fields("ON-LOAN") = True
            .Update
        End With
        ' Update TOTAL ITEM BORROWED by increment 1
        SQL = "SELECT TOT_ITEM_BORROWED FROM PATRON WHERE PATRON_NUM = '" _
            & txtPatronNum & "'"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        With gLibDyna
            .Edit
                .Fields("TOT_ITEM_BORROWED") = _
                    CInt(.Fields("TOT_ITEM_BORROWED")) + 1
            .Update
        End With
        
        ' Update Active Status
        SQL = "SELECT PATRON.TOT_ITEM_BORROWED, " _
            & "[MEMBER-TYPE].MAX_BORROWING, PATRON.ACTIVE_STATUS " _
            & "FROM [MEMBER-TYPE] INNER JOIN PATRON ON " _
            & "[MEMBER-TYPE].MEMBER_CODE = PATRON.MEMBER_CODE WHERE " _
            & "PATRON_NUM = '" & txtPatronNum & "'"
        Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        If gLibDyna.Fields("MAX_BORROWING") = _
            gLibDyna.Fields("TOT_ITEM_BORROWED") Then
                gLibDyna.Edit
                gLibDyna.Fields("ACTIVE_STATUS") = False
                gLibDyna.Update
        End If
    
    ' Update display - Should use Add
    fillLvwLoan txtPatronNum
    fillTvw
    
    tPrevious = txtPatronNum
    
'    txtPatronNum = ""
'    txtAcqNum = ""
    
    dbRetrieveMaterialLoan
        
Exit Sub
Error_Handler:
    DisplayError "Database Error"
End Sub


Private Sub Form_Load()
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'    InitDB
    CenterForm Me, frmMDIMainMenu
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    
    fillTvw
    fillLvwPatronBorrow
    gtTableActive = "LOAN"
    dbRetrieve lstPatronNum, "PATRON", "PATRON_NUM", 1
'    dbRetrieve lstAcqNum, "MATERIAL", "ACQUISITION_NUM", 1
    dbRetrieveMaterialLoan
    frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = "Double click Item to select"
    fLoanEnabled = True
End Sub

Private Sub dbRetrieveMaterialLoan()
    lstAcqNum.Clear
    SQL = "SELECT DISTINCT MATERIAL.ACQUISITION_NUM FROM MATERIAL WHERE [ON-LOAN] = " & False & " AND EXISTENCE = " & True
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    While Not gLibSnap.EOF
        If Not IsNull(gLibSnap.Fields(0)) Then
            lstAcqNum.AddItem gLibSnap.Fields(0)
        End If
        gLibSnap.MoveNext
    Wend
    If lstAcqNum.ListCount Then lstAcqNum.ListIndex = 0
End Sub

Private Sub lblMaterialTitle_Change()
    If lblMaterialTitle = "" Then
        fUpdateLoanA = False
    Else
        fUpdateLoanA = True
    End If
    
    If fUpdateLoanA And fUpdateLoanP Then
        cmd3LoanUpdate.Enabled = True
    Else
        cmd3LoanUpdate.Enabled = False
    End If
End Sub

Private Sub lblPatronName_Change()
    If lblPatronName = "" Then
        fUpdateLoanP = False
    Else
        CheckActiveStatus txtPatronNum
        fUpdateLoanP = True
        tPrevious = txtPatronNum
    End If
    If fUpdateLoanA And fUpdateLoanP Then
        cmd3LoanUpdate.Enabled = True
    Else
        cmd3LoanUpdate.Enabled = False
    End If
End Sub


Private Sub lstAcqNum_DblClick()
    txtAcqNum = lstAcqNum
End Sub

Private Sub lstPatronNum_DblClick()
    fillLvwLoan lstPatronNum
    DoEvents
    txtPatronNum = lstPatronNum
End Sub

Private Sub lvwLoan_DblClick()
    fillLvwLoan tSelectedItem
    txtPatronNum = tSelectedItem
End Sub

Private Sub lvwLoan_ItemClick(ByVal Item As ListItem)
    tSelectedItem = CStr(Item)
End Sub

Private Sub tvwLibrary_NodeClick(ByVal Node As Node)
    If Parse(Node.FullPath, 2, "\") <> "" Then
        txtPatronNum = Parse(Node.FullPath, 2, "\")
        fillLvwLoan CStr(Parse(Node.FullPath, 2, "\"))
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


Private Sub txtAcqNum_GotFocus()
    lblIndicator = "Acquisition Number ..."
End Sub

Private Sub txtAcqNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmd3LoanUpdate.Enabled = True Then cmd3LoanUpdate.SetFocus
    End If
End Sub

Private Sub txtAcqNum_LostFocus()
    lblIndicator = ""
End Sub

Private Sub txtPatronNum_Change()
    On Local Error GoTo Error_Handler:
    
    srchLikeLstItem Me, lstPatronNum, txtPatronNum
    
    If Len(txtPatronNum) Then
        SQL = "SELECT NAME FROM PATRON WHERE PATRON_NUM = '" _
            & txtPatronNum & "'"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If gLibSnap.EOF Then
            lblPatronName = ""
        Else
            lblPatronName = gLibSnap.Fields("NAME")
            fillLvwLoan txtPatronNum
'            CheckActiveStatus txtPatronNum
        End If
    ElseIf Not (Len(txtPatronNum)) Then ' Blank entry
        lblPatronName = ""
    End If
    
Exit Sub

Error_Handler:
    DisplayError ""
End Sub


Private Sub CheckActiveStatus(ptPatronNum As String)
''Debug.Print Not (Len(ptPatronNum))
    If Len(ptPatronNum) = 0 Then Exit Sub
    
    Dim fClearedFine As Boolean
    SQL = "SELECT TOT_ITEM_BORROWED, ACTIVE_STATUS " _
        & "FROM PATRON WHERE PATRON_NUM = '" & ptPatronNum & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    If (gLibSnap.Fields("ACTIVE_STATUS") = False) Then
        
        ' Check Total Item Borrowed
        SQL = "SELECT PATRON.TOT_ITEM_BORROWED, [MEMBER-TYPE].MAX_BORROWING " _
            & "FROM [MEMBER-TYPE] INNER JOIN PATRON ON " _
            & "[MEMBER-TYPE].MEMBER_CODE = PATRON.MEMBER_CODE WHERE " _
            & "PATRON_NUM = '" & ptPatronNum & "'"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        
        If gLibSnap.Fields("TOT_ITEM_BORROWED") = _
            gLibSnap.Fields("MAX_BORROWING") Then
                MsgBox "Over Limit", vbCritical, "Maximum Borrowing"
                fLoanEnabled = False
                txtPatronNum = ""
                txtPatronNum.SetFocus
        End If
        
        ' Check Cleared Fine
        SQL = "SELECT LOAN.DATE_RETURN, LOAN.CLEARED_FINE FROM LOAN " _
            & "WHERE LOAN.PATRON_NUM = '" & ptPatronNum & "'"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If Not (gLibSnap.EOF) Then
            fClearedFine = True     ' Init
            With gLibSnap
            While Not .EOF
                ' Check if any fine for a loan is uncleared for material
                '   that has been returned but not paid for the fine
                If Not (IsNull((.Fields("DATE_RETURN")))) And _
                    (.Fields("CLEARED_FINE") = False) Then _
                    fClearedFine = False
                .MoveNext
            Wend
            End With
        Else
            fClearedFine = True
        End If
    
        If Not (fClearedFine) Then
            MsgBox "Please clear debt", vbInformation
'            txtPatronNum = ""
            txtPatronNum.SetFocus
            fLoanEnabled = False
            
        End If
    End If

End Sub


Private Sub fillTvw()
    Dim colNode As Node
    tvwLibrary.Nodes.Clear
'    SQL = "SELECT DISTINCT PATRON_NUM FROM LOAN ORDER BY PATRON_NUM ASC"
    SQL = "SELECT DISTINCT LOAN.PATRON_NUM, PATRON.TOT_ITEM_BORROWED FROM PATRON INNER JOIN LOAN ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE TOT_ITEM_BORROWED > 0 ORDER BY LOAN.PATRON_NUM ASC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    Set colNode = tvwLibrary.Nodes.Add(, , "Root", "Patron", 1)
    While Not gLibSnap.EOF
        Set colNode = tvwLibrary.Nodes.Add("Root", tvwChild, "C" _
            & CStr(gLibSnap.Fields("PATRON_NUM")), CStr(gLibSnap.Fields("PATRON_NUM")), 2)
        gLibSnap.MoveNext
        colNode.EnsureVisible
    Wend
End Sub


Private Sub txtPatronNum_GotFocus()
    lblIndicator = "Patron Number ..."
End Sub

Private Sub txtPatronNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAcqNum.SetFocus
    End If
End Sub

Private Sub txtPatronNum_LostFocus()
    lblIndicator = ""
End Sub


'***************************************************************************

Private Sub fillLvwLoan(tPatronNum As String)
    Dim colColumn As ColumnHeader
    Dim colItem As ListItem
    Dim iIndex As Integer
    clearLvw
    SQL = "SELECT PATRON.PATRON_NUM, LOAN.ACQUISITION_NUM, MATERIAL.TITLE, " & _
        "LOAN.DATE_BORROW " & _
        "FROM PATRON INNER JOIN (MATERIAL INNER JOIN LOAN ON " & _
        "MATERIAL.ACQUISITION_NUM = LOAN.ACQUISITION_NUM) ON " & _
        "PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE LOAN.PATRON_NUM " & _
        "= '" & tPatronNum & "'" & " AND LOAN.CLEARED_FINE = " & False

    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)

    ' Display Not Returned Material or Cleared_Fine = False
    Set colColumn = lvwLoan.ColumnHeaders.Add(, "ACQUISITION NO.", "ACQUISITION NO.", lvwLoan.Width / 3)
    Set colColumn = lvwLoan.ColumnHeaders.Add(, "MATERIAL TITLE", "MATERIAL TITLE", lvwLoan.Width / 3)
    Set colColumn = lvwLoan.ColumnHeaders.Add(, "DATE BORROW", "DATE BORROW", lvwLoan.Width / 3)
    
    While Not gLibSnap.EOF
        Set colItem = lvwLoan.ListItems.Add(, , CStr(gLibSnap.Fields("ACQUISITION_NUM")))
        sAcqNum = gLibSnap.Fields("ACQUISITION_NUM")
        If Not IsNull(gLibSnap.Fields("TITLE")) Then
            colItem.SubItems(1) = CStr(gLibSnap.Fields("TITLE"))
        End If
        If Not IsNull(gLibSnap.Fields("DATE_BORROW")) Then
            colItem.SubItems(2) = CStr(gLibSnap.Fields("DATE_BORROW"))
        End If
        gLibSnap.MoveNext
    Wend
    fLoanEnabled = True
End Sub

Private Sub fillLvwPatronBorrow()
    Dim colColumn As ColumnHeader
    Dim colItem As ListItem
    Dim iIndex As Integer
    clearLvw
    lvwLoan.View = lvwIcon
    
    SQL = "SELECT DISTINCT LOAN.PATRON_NUM, PATRON.TOT_ITEM_BORROWED FROM PATRON INNER JOIN LOAN ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE TOT_ITEM_BORROWED > 0 ORDER BY LOAN.PATRON_NUM ASC"
    
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)

    ' Display Not Returned Material or Cleared_Fine = False
    Set colColumn = lvwLoan.ColumnHeaders.Add(, "PATRON NO.", "PATRON NO.", lvwLoan.Width / 3)
    While Not gLibSnap.EOF
        Set colItem = lvwLoan.ListItems.Add(, , CStr(gLibSnap.Fields("PATRON_NUM")), 2)
        gLibSnap.MoveNext
    Wend
    AdjustColumnWidth lvwLoan, True
End Sub

' Clear ListView control and set view to viewReport
Private Sub clearLvw()
  lvwLoan.View = lvwReport
  lvwLoan.ColumnHeaders.Clear
  lvwLoan.ListItems.Clear
End Sub

