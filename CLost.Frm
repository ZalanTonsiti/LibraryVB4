VERSION 4.00
Begin VB.Form frmReturnPatronNum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please enter Acquisition Number..."
   ClientHeight    =   3045
   ClientLeft      =   1050
   ClientTop       =   3210
   ClientWidth     =   9555
   Height          =   3450
   Icon            =   "CLost.frx":0000
   Left            =   990
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   9555
   Top             =   2865
   Width           =   9675
   Begin VB.OptionButton optLostReason 
      Caption         =   "Lost &Undetected"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   2520
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optLostReason 
      Caption         =   "Lost by &Patron"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame fraLostMaterial 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9375
      Begin VB.ListBox lstAcqNum 
         Height          =   1230
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtAcqNum 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.ListBox lstPatronNum 
         Height          =   1230
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtPatronNum 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblInoAcqNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition Number"
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMaterialTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   11
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label lblInfoAcqName 
         Alignment       =   1  'Right Justify
         Caption         =   "Material Title"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblInfoPatronNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Number"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblInfoPatronName 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Name"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblPatronName 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
      End
   End
   Begin Threed.SSCommand cmd3OKCancel 
      Height          =   300
      Index           =   0
      Left            =   6960
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&OK"
      BevelWidth      =   1
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmd3OKCancel 
      Height          =   300
      Index           =   1
      Left            =   8280
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Close"
      BevelWidth      =   1
      Outline         =   0   'False
   End
End
Attribute VB_Name = "frmReturnPatronNum"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmReturnPatronNum
'***************************************************************************

Private Sub Lost(fLostStatus As Boolean)
    Dim tPatronNum As String
    Dim cTotalFine As Currency
    Dim cBookPrice As Currency
    Dim xDateBorrow As Date
    Dim iBorrowDuration As Integer
    Dim sFineRate As Single
    ' Check Loan Details
    SQL = "SELECT PATRON_NUM, DATE_RETURN FROM LOAN WHERE ACQUISITION_NUM = '" & txtAcqNum & "' ORDER BY DATE_BORROW DESC"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    tPatronNum = gLibDyna.Fields("PATRON_NUM")
    If Not gLibDyna.EOF Then
        If IsNull(gLibDyna.Fields("DATE_RETURN")) Then
            ' Update DATE RETURN of lost material
            With gLibDyna
                .Edit
                    .Fields("DATE_RETURN") = Date
                .Update
            End With
            ' Update MATERIAL EXISTENCE status
            SQL = "SELECT [ON-LOAN], EXISTENCE FROM MATERIAL WHERE ACQUISITION_NUM = '" & txtAcqNum & "'"
            Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            With gLibDyna
                .Edit
                    .Fields("EXISTENCE") = False
                    .Fields("ON-LOAN") = False
                .Update
            End With
            ' Update PATRON TABLE
            SQL = "SELECT ACTIVE_STATUS, TOT_ITEM_BORROWED FROM PATRON WHERE PATRON_NUM = '" & tPatronNum & "'"
            Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            With gLibDyna
            .Edit
                ' The default value of ACTIVE_STATUS of LOAN is set to FALSE
                '.Fields("ACTIVE_STATUS") = False
                .Fields("TOT_ITEM_BORROWED") = CInt(.Fields("TOT_ITEM_BORROWED")) - 1
            .Update
            End With
            ' Calculate Total Fine
                ' Check Book Price
            SQL = "SELECT SUPPLY.COST FROM MATERIAL" & _
                  " INNER JOIN SUPPLY ON MATERIAL.ACQUISITION_NUM = SUPPLY.ACQUISITION_NUM " & _
                  "WHERE SUPPLY.ACQUISITION_NUM = '" & txtAcqNum & "'"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            cBookPrice = gLibSnap.Fields("COST")
            SQL = "SELECT PATRON.PATRON_NUM, [MATERIAL-TYPE].DURATION_OF_BORROW, " & _
                  "LOAN.ACQUISITION_NUM, LOAN.DATE_BORROW, LOAN.DATE_RETURN, " & _
                  "[MEMBER-TYPE].FINE_RATE_PERDAY FROM ([MEMBER-TYPE] INNER JOIN " & _
                  "PATRON ON [MEMBER-TYPE].MEMBER_CODE = PATRON.MEMBER_CODE) INNER JOIN " & _
                  "(([MATERIAL-TYPE] INNER JOIN MATERIAL ON [MATERIAL-TYPE].MATERIAL_CODE " & _
                  "= MATERIAL.MATERIAL_CODE) INNER JOIN LOAN ON MATERIAL.ACQUISITION_NUM = " & _
                  "LOAN.ACQUISITION_NUM) ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE " & _
                  "LOAN.ACQUISITION_NUM = '" & txtAcqNum & "' AND LOAN.PATRON_NUM = '" & _
                  tPatronNum & "' AND LOAN.DATE_RETURN = #" & Date & "# ORDER BY DATE_RETURN DESC;"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            iBorrowDuration = gLibSnap.Fields("DURATION_OF_BORROW")
            xDateBorrow = gLibSnap.Fields("DATE_BORROW")
            sFineRate = gLibSnap.Fields("FINE_RATE_PERDAY")
            
            ' Check if the lost material is late
            If xDateBorrow + iBorrowDuration < Date Then
              If fLostStatus Then
                cTotalFine = ((Date - (xDateBorrow + iBorrowDuration)) * sFineRate) + cBookPrice
              Else
                cTotalFine = ((Date - (xDateBorrow + iBorrowDuration)) * sFineRate)
              End If
            Else
                cTotalFine = cBookPrice
            End If
            SQL = "SELECT LOAN.TOT_FINE, LOAN.CLEARED_FINE FROM LOAN WHERE ACQUISITION_NUM = '" & _
                  txtAcqNum & "' AND PATRON_NUM = '" & tPatronNum & _
                  "' AND DATE_RETURN = #" & Date & "# ORDER BY DATE_RETURN DESC"
            Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            With gLibDyna
                .Edit
                    .Fields("TOT_FINE") = CSng(cTotalFine)
                .Update
                frmReturn.fillLvwReturn txtPatronNum
            End With
        Else
            MsgBox "The material has been returned.", vbInformation
        End If
    Else        ' Record not found
        MsgBox "No record of borrowing found", vbExclamation
        Exit Sub
End If
End Sub



Private Sub cmd3OKCancel_Click(Index As Integer)
    
    
    Select Case Index
        Case 0
            
            If (Not CBool(Len(lblPatronName)) Or Not CBool(Len(lblMaterialTitle))) Then
                MsgBox "Please fill both entries with valid information", vbInformation, strSystem
                Exit Sub
            End If
            
            Dim tConf$
            Dim tReason$
            Dim vResponse
            tConf = "Please confirm the following :-" & vbCrLf & vbCrLf
            tConf = tConf & "Material Acquisition Number : " & txtAcqNum & vbCrLf
            tConf = tConf & "Material Name                      : " & lblMaterialTitle & vbCrLf & vbCrLf
            tConf = tConf & "Patron Number                     : " & txtPatronNum & vbCrLf
            tConf = tConf & "Patron Name                        : " & lblPatronName & vbCrLf & vbCrLf
            If optLostReason(0).Value Then
                tReason$ = mReplaceCharacter("&", "", optLostReason(0).Caption)
                tConf = tConf & tReason$
            Else
                tReason$ = mReplaceCharacter("&", "", optLostReason(1).Caption)
                tConf = tConf & tReason$
            End If
                vResponse = MsgBox(tConf$, vbYesNo, strsytem & " - Confirmation")
                If vResponse = vbYes Then
                    If optLostReason(0).Value Then
                        Lost 1
                    Else
                        Lost 0
                    End If
                End If
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    dbRetrieveMaterialLoan
'    dbRetrieve lstPatronNum, "LOAN", "PATRON_NUM"
    dbRetrievePatronLoan
End Sub


Private Sub dbRetrievePatronLoan()
    
    SQL = "SELECT DISTINCT LOAN.PATRON_NUM, PATRON.TOT_ITEM_BORROWED FROM PATRON INNER JOIN LOAN ON PATRON.PATRON_NUM = LOAN.PATRON_NUM WHERE TOT_ITEM_BORROWED > 0 ORDER BY LOAN.PATRON_NUM ASC"
    
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    ' Display Not Returned Material or Cleared_Fine = False
    While Not gLibSnap.EOF
        lstPatronNum.AddItem gLibSnap.Fields("PATRON_NUM")
        gLibSnap.MoveNext
    Wend

End Sub

Private Sub dbRetrieveMaterialLoan()
    lstAcqNum.Clear
    SQL = "SELECT MATERIAL.ACQUISITION_NUM FROM MATERIAL WHERE [ON-LOAN] = " & True
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    With gLibSnap
        While Not .EOF
            If Not IsNull(.Fields(0)) Then
                lstAcqNum.AddItem .Fields(0)
            End If
            .MoveNext
        Wend
    End With
    
    If lstAcqNum.ListCount Then lstAcqNum.ListIndex = 0
End Sub

Private Sub lblMaterialTitle_Change()
    If lblMaterialTitle <> "" Then
'        tCurrentAcquisition = txtAcqNum
        SQL = "SELECT PATRON_NUM FROM LOAN WHERE ACQUISITION_NUM = '" & txtAcqNum & _
            "' ORDER BY DATE_BORROW DESC"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        
        If Not gLibSnap.EOF Then
            srchLstItem Me, lstPatronNum, CStr(gLibSnap.Fields(0))
            lstPatronNum_DblClick
        End If
    End If

End Sub

Private Sub lblPatronName_Change()
    If Len(lblPatronName) Then
        dbShowCorrMaterial
    Else
        dbRetrieveMaterialLoan
    End If
    txtAcqNum = ""
End Sub

Private Sub dbShowCorrMaterial()
    lstAcqNum.Clear
    SQL = "SELECT ACQUISITION_NUM, CLEARED_FINE, DATE_RETURN FROM LOAN WHERE PATRON_NUM = '" & _
        txtPatronNum & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    While Not gLibSnap.EOF
      If IsNull(gLibSnap.Fields("DATE_RETURN")) Then    ' Filter unReturnMaterial Only
        If Not IsNull(gLibSnap.Fields(0)) Then
            lstAcqNum.AddItem gLibSnap.Fields(0)
        End If
      End If
      gLibSnap.MoveNext
    Wend
    If lstAcqNum.ListCount Then
        lstAcqNum.ListIndex = 0
    Else
    
    End If
End Sub

Private Sub lstAcqNum_DblClick()
    txtAcqNum = lstAcqNum
End Sub

Private Sub lstPatronNum_DblClick()
    txtPatronNum = lstPatronNum
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
'    On Local Error GoTo Error_Handler:
    
    srchLikeLstItem Me, lstPatronNum, txtPatronNum
    
    If Len(txtPatronNum) Then
        SQL = "SELECT NAME FROM PATRON WHERE PATRON_NUM = '" _
            & txtPatronNum & "'"
        Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
        If gLibSnap.EOF Then
            lblPatronName = ""
        Else
            If fSrchLstItem(Me, lstPatronNum, txtPatronNum) Then
                lblPatronName = gLibSnap.Fields("NAME")
    '            fillLvwLoan txtPatronNum
'                CheckActiveStatus txtPatronNum
            Else
                lblPatronName = ""
            End If
        End If
    ElseIf Not (Len(txtPatronNum)) Then ' Blank entry
        lblPatronName = ""
    End If
End Sub
