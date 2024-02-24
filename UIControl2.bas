Attribute VB_Name = "mUIControl"
'***************************************************************************
' mUIControl - General UI Control
' Update ' 18/02/98
'***************************************************************************


Public Sub CtrCommand3D(pCtrFrm As Form, Index As Integer)
On Error GoTo Error_Handler
    Select Case Index
        Case 0                              ' New
            gfNewStatus = True
            CtrDB pCtrFrm, 0
            pCtrFrm.lvwLibrary.Enabled = 0
        Case 1                              ' Edit
            If gtCurrentIndex = "" Then
                MsgBox "Please click a record "
                Exit Sub
            End If
            gfNewStatus = False
            CtrDB pCtrFrm, 1
            pCtrFrm.lvwLibrary.Enabled = 0
        Case 2                              ' Enter
            If gfNewStatus Then
                Set gLibDyna = gLibDB.OpenRecordset(gtTableActive, dbOpenDynaset)
                gLibDyna.AddNew
                AssignFieldText pCtrFrm, 1
            Else
                SQL = "SELECT * FROM [" & gtTableActive & "] WHERE [" & gtTableIndex & "] = '" & gtCurrentIndex & "'"
                Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
                gLibDyna.Edit
                AssignFieldText pCtrFrm
            End If
            gLibDyna.Update
            fillLvw pCtrFrm, 5
            CtrDB pCtrFrm, 2
            pCtrFrm.lvwLibrary.Enabled = 1

        Case 3  ' Cancel
           CtrDB pCtrFrm, 3
            pCtrFrm.lvwLibrary.Enabled = 1

    End Select
Exit Sub
Error_Handler:
    DisplayError ""



End Sub


Public Sub CtrDB(frm As Form, pIndex As Integer)
On Error GoTo Error_Handler
    
    Select Case pIndex
        Case 0  ' New
            
            frm.cmd3DBCtrl(0).Enabled = False
            frm.cmd3DBCtrl(1).Enabled = False
            frm.cmd3DBCtrl(2).Enabled = True
            frm.cmd3DBCtrl(3).Enabled = True
            FieldControlNew frm
        
        Case 1  ' Edit
            
            frm.cmd3DBCtrl(0).Enabled = False
            frm.cmd3DBCtrl(1).Enabled = False
            frm.cmd3DBCtrl(2).Enabled = True
            frm.cmd3DBCtrl(3).Enabled = True
            FieldControlEdit frm
        
        Case 2, 3 ' Enter or Cancel
            
            frm.cmd3DBCtrl(2).Enabled = False
            frm.cmd3DBCtrl(3).Enabled = False
            frm.cmd3DBCtrl(0).Enabled = True
            frm.cmd3DBCtrl(1).Enabled = True
            
            FieldControlDisabled frm
    End Select
Exit Sub
Error_Handler:
    DisplayError ""
    
End Sub

Public Sub FieldControlEdit(frm As Form)
    Dim Index As Integer
    
    giFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    ' Started from 2 because the Primary Key (usually first Field) cannot be deleted or changed
    '   of REFERENTIAL INTEGERITY Err#3200
    For Index = 2 To giFldTotal%
        frm.Controls(Index).Enabled = True
    Next
End Sub


Public Sub FieldControlNew(frm As Form)
    Dim Index As Integer
    
   giFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    
    For Index = 1 To giFldTotal% 'frm.Controls.Count - 1
        If (TypeOf frm.Controls(Index) Is TextBox Or TypeOf frm.Controls(Index) Is ComboBox) Then
            frm.Controls(Index).Enabled = True
            frm.Controls(Index).Text = ""
        ElseIf (TypeOf frm.Controls(Index) Is CheckBox) Then
            frm.Controls(Index).Enabled = True
            frm.Controls(Index).Value = 0
        End If
    Next
    ' Find Default Value for fields and assign to controls
End Sub

'   Set TextBox and CheckBox to DISABLED
Public Sub FieldControlDisabled(frm As Form)
    Dim Index As Integer
    
    giFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    
    For Index = 1 To giFldTotal% 'frm.Controls.Count - 1
        frm.Controls(Index).Enabled = False
    Next
End Sub

'   Clear TextBox and CheckBox
Public Sub FieldControlClear(frm As Form)
    Dim Index As Integer
    For Index = 0 To frm.Controls.Count - 1
        If (TypeOf frm.Controls(Index) Is TextBox) Then
            frm.Controls(Index).Text = ""
        ElseIf (TypeOf frm.Controls(Index) Is CheckBox) Then
            frm.Controls(Index).Value = 0
        Else
        End If
    Next
End Sub


' Clear ListView control and set view to viewReport
Public Sub clearLvw(pCtrLvw As Control)
    Dim clmX                As ColumnHeader
  pCtrLvw.View = lvwReport
  pCtrLvw.ColumnHeaders.Clear
  pCtrLvw.ListItems.Clear
'If gfSearchMode Then
'        Set clmX = pCtrLvw.ColumnHeaders.Add(, gtLibErr(12), gtLibErr(12))
'        ' Set General Option; Show All Entry detail or Optimized
'        AdjustColumnWidth pCtrLvw, True
'End If
End Sub


'===========================================================================
' DESC:     Fill ColumnHeader with Field content
' PARAMS:   frm
' PROJECT NOTES : Many ListView Many Tables
'===========================================================================
Public Sub fillLvw(frm As Form, iDiv As Integer, Optional vSql)
On Local Error GoTo Error_Handler
    MsgBar "Retrieving records", True
        
    clearLvw frm.lvwLibrary
    Dim xStart As Date
    Dim clmX                As ColumnHeader
    Dim itmX                As ListItem
    Dim ftfldvalue          As String
    xStart = Now
    Screen.MousePointer = vbHourglass
    
    If IsMissing(vSql) Then
        Set gLibSnap = gLibDB.OpenRecordset(gtTableActive, dbOpenSnapshot) '
    Else
        Set gLibSnap = gLibDB.OpenRecordset(vSql, dbOpenSnapshot)
    End If
    
If Not gLibSnap.EOF Then
    clearLvw frm.lvwLibrary
    
    ' Extract number of fields in a table. Minus 1 because it start from 1
    giFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count - 1
    
    ' Fill ColumnHeader of ListItems control with fields name
    For fiIndex% = 0 To giFldTotal% '- 1
        ' Extract name of fields name
        ftFldName$ = mReplaceCharacter("_", " ", CStr(gLibDB.TableDefs(gtTableActive).Fields(fiIndex%).Name))
        ' Set ColumnHeaders of ListView Control with fields name
        Set clmX = frm.lvwLibrary.ColumnHeaders.Add(, ftFldName$, ftFldName$, _
                                           (frm.lvwLibrary.Width / iDiv))
    Next fiIndex%
    
    
    ' Fill ListItems and SubItems of ListView control with Database content
    While Not gLibSnap.EOF
        ' Set ListItems of Listview control with fields value
        ftfldvalue = CStr(gLibSnap.Fields(CStr(gLibDB.TableDefs(gtTableActive).Fields(0).Name)))
        If Not IsNull(ftfldvalue) Then Set itmX = _
            frm.lvwLibrary.ListItems.Add(, , CStr(ftfldvalue))
        ' Set an icon from ImageList1
        itmX.Icon = 1
        ' Set an icon from ImageList2
        itmX.SmallIcon = 1
        For fiIndex% = 1 To giFldTotal%
            ftFldName$ = CStr(gLibDB.TableDefs(gtTableActive).Fields(fiIndex%).Name)
            If Not IsNull(gLibSnap.Fields(ftFldName$)) Then
                ' Set SubItems of Listview control with fields value
                ftfldvalue = CStr(gLibSnap.Fields(ftFldName$))
                If ftfldvalue = "True" Or ftfldvalue = "False" Then
                    ftfldvalue = gtConvBoolString(ftfldvalue)
                End If
                itmX.SubItems(fiIndex%) = ftfldvalue
            End If
        Next fiIndex%
        ' Move to next record.
        gLibSnap.MoveNext
    Wend
    
    ' Show Content or Header
    If getLibSet("ShowHeader") = "True" Then
        AdjustColumnWidth frm.lvwLibrary, False
    Else
        AdjustColumnWidth frm.lvwLibrary, True
    End If
Else
    ' Show message no item to be displayed
    Set clmX = frm.lvwLibrary.ColumnHeaders.Add(, gtLibErr(12), gtLibErr(12))
    ' Set General Option; Show All Entry detail or Optimized
    AdjustColumnWidth frm.lvwLibrary, True
End If
    Screen.MousePointer = vbDefault
    MsgBar "", False

' BenchMark
' frmMDIMainMenu.sbrLibrary.Panels.Item(2).Text = DateDiff("s", xStart, Now)
' MsgBox DateDiff("s", xStart, Now)
GoTo End_Sub
Error_Handler:
    DisplayError ""
    Resume End_Sub
End_Sub:
End Sub


Public Sub fillCurrentDetail(pCtrFrm As Form, tID As String)
    Dim Index    As Integer
    Dim iFldTotal%
    
    ' Clear Current TextBox
    FieldControlClear pCtrFrm
    
    ' Extract number of fields in a table. Minus 1 because it start from 1
    iFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count - 1
    
    SQL = "SELECT * FROM [" & gtTableActive & "] WHERE [" & gtTableIndex & "] = '" & tID & "'"
    
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    If Not gLibSnap.EOF Then
        With gLibSnap
            For Index = 0 To iFldTotal%
                If Not IsNull(.Fields(Index)) Then
                    If TypeOf pCtrFrm.Controls(Index + 1) Is TextBox Then
                       pCtrFrm.Controls(Index + 1).Text = .Fields(Index)
                    ElseIf TypeOf pCtrFrm.Controls(Index + 1) Is CheckBox Then
                       pCtrFrm.Controls(Index + 1).Value = Abs(.Fields(Index))
                    ElseIf TypeOf pCtrFrm.Controls(Index + 1) Is ComboBox Then
'                       pctrFrm.Controls(Index + 1).Text = .Fields(Index)
                        'NEEDFIX
                        srchLstItem pCtrFrm, pCtrFrm.Controls(Index + 1), .Fields(Index)
                    End If
                End If
            Next
        End With
    End If
End Sub


Public Sub AssignFieldText(pCtrFrm As Form, Optional vOptStart As Variant)
    Dim iIndexStart
    If IsMissing(vOptStart) Then
        iIndexStart = 2
    Else
        iIndexStart = 1
    End If
    iFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    ' Start from 2 Exclude the Primary Key
    For Index = iIndexStart To iFldTotal%
        If pCtrFrm.Controls(Index) <> "" Then
            gLibDyna.Fields(Index - 1) = pCtrFrm.Controls(Index) 'CStr(pctrFrm.Controls(Index))
        End If
    Next Index
End Sub

Public Sub BuildSearchSQL(pCtrFrm As Form)
    
    Dim tQuery          As String
    Dim tString         As String
    Set gLibSnap = gLibDB.OpenRecordset(gtTableActive, dbOpenDynaset)
    iFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    Debug.Print iFldTotal%, "iFldTotal%"
    tQuery = "SELECT * FROM [" & gtTableActive & "] WHERE "
    
    For Index = 0 To iFldTotal% - 1
      If pCtrFrm.Controls(Index + 1).Tag = "Dirty" Then ' Build the clause for Dirtied Field Only
        If TypeOf pCtrFrm.Controls(Index + 1) Is TextBox Then
            tString = pCtrFrm.Controls(Index + 1).Text
        ElseIf TypeOf pCtrFrm.Controls(Index + 1) Is CheckBox Then
            tString = Abs(pCtrFrm.Controls(Index + 1).Value)
        ElseIf TypeOf pCtrFrm.Controls(Index + 1) Is ComboBox Then
            tString = pCtrFrm.Controls(Index + 1)
        End If
        
        ' Determine the DataType of Field to assign appropriate SQL clause
        If gLibSnap.Fields(Index).Type = 6 Then     ' Tot_Item_Borrowed
            If tString = "" Then tString = "0"
            tQuery = tQuery & "[" & gLibSnap.Fields(Index).Name & _
                "]" & " >= " & tString & " AND "     ' NEED FIX - Use of the > = <
        
        
        ElseIf gLibSnap.Fields(Index).Type = 5 Then     ' Currency
            If tString = "" Then tString = "0"
            tQuery = tQuery & "[" & gLibSnap.Fields(Index).Name & _
                "]" & " >= " & tString & " AND "     ' NEED FIX - Use of the > = <
        
        ElseIf gLibSnap.Fields(Index).Type = 8 Then     ' Date
            If tString = "" Then tString = "1/1/80"
            tQuery = tQuery & "[" & gLibSnap.Fields(Index).Name & _
                "]" & " >= #" & tString & "# AND "     ' NEED FIX - Use of the > = <
        
        ElseIf gLibSnap.Fields(Index).Type = 1 Then 'Boolean
            tQuery = tQuery & "[" & gLibSnap.Fields(Index).Name & _
                "]" & " = " & CBool(tString) & " AND "
        
        ElseIf gLibSnap.Fields(Index).Type = 10 Then
            tQuery = tQuery & "[" & gLibSnap.Fields(Index).Name & _
                "]" & " LIKE '*" & tString & "*' AND "
        End If
      End If
    Next Index
    tQuery = Left(tQuery, Len(tQuery) - 4)
    fillLvw pCtrFrm, 2, tQuery
End Sub

Public Sub ShowFirst(pCtrFrm As Form)
    SQL = "SELECT [" & gtTableIndex & "] FROM [" & gtTableActive & "] ORDER BY [" & gtTableIndex & "] ASC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    ' Fill Detail of member if table is not null
    If Not gLibSnap.EOF Then fillCurrentDetail pCtrFrm, CStr(gLibSnap.Fields(0))
End Sub


'===========================================================================
'   DESC    :   Search through ListIndex Property of Object, with compare of
'               a search string
'   Apply To:   ComboBox, DirListBox, DriveListBox, FileListBox, ListBox
'===========================================================================
Public Sub srchLstItem(pCtrFrm As Form, objLst As Control, ptSrch$)
    If ptSrch$ = "" Then
        objLst.ListIndex = 0
        Exit Sub
    End If
    For fiIndex% = 0 To objLst.ListCount - 1
        If StrComp(ptSrch$, objLst.List(fiIndex%), 0) = 0 Then
            objLst.ListIndex = fiIndex%
            Exit Sub
        End If
    Next fiIndex%
            objLst.ListIndex = objLst.ListCount - 1 'Not Found
End Sub


Public Sub CtrDisableLvwView()
    frmMDIMainMenu.tbrLibrary.Buttons("Find").Enabled = False
    frmMDIMainMenu.mnuViewElip(0).Enabled = False
    frmMDIMainMenu.mnuViewElip(1).Enabled = False
    frmMDIMainMenu.mnuViewElip(2).Enabled = False
    frmMDIMainMenu.mnuViewElip(3).Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Value = tbrUnpressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Value = tbrUnpressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Value = tbrUnpressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Value = tbrUnpressed

End Sub


Public Sub CtrDisableLvwViewMain()
    frmMDIMainMenu.tbrLibrary.Buttons("Find").Enabled = False
    frmMDIMainMenu.mnuViewElip(0).Enabled = False
    frmMDIMainMenu.mnuViewElip(1).Enabled = False
    frmMDIMainMenu.mnuViewElip(2).Enabled = False
    frmMDIMainMenu.mnuViewElip(3).Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Enabled = True
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Enabled = True
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Enabled = True
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Enabled = False
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Value = tbrPressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Value = tbrUnpressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Value = tbrUnpressed
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Value = tbrUnpressed
End Sub


Public Sub CtrEnableLvwViewMain()
    On Local Error GoTo Error_Handler
    frmMDIMainMenu.mnuViewElip(0).Enabled = True
    frmMDIMainMenu.mnuViewElip(1).Enabled = True
    frmMDIMainMenu.mnuViewElip(2).Enabled = True
    If gtTableIndex = "" Then
        frmMDIMainMenu.mnuViewElip(3).Enabled = False
    Else
        frmMDIMainMenu.mnuViewElip(3).Enabled = True
    End If
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Enabled = True
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Enabled = True
    frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Enabled = True
    If gtTableIndex = "" Then
        gtCurrentIndex = ""
        frmMDIMainMenu.tbrLibrary.Buttons("Find").Enabled = False
        frmMDIMainMenu.tbrLibrary.Buttons("Del").Enabled = False
        frmMDIMainMenu.tbrLibrary.Buttons("New").Enabled = False
        frmMDIMainMenu.tbrLibrary.Buttons("Print").Enabled = False
        frmMDIMainMenu.mnuViewElip(0).Checked = True
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Value = tbrPressed
        frmMDIMainMenu.tbrLibrary.Buttons("Find").Value = tbrUnpressed
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Value = tbrUnpressed
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Enabled = False
    Else
        frmMDIMainMenu.tbrLibrary.Buttons("Find").Enabled = True
        frmMDIMainMenu.tbrLibrary.Buttons("Del").Enabled = True
        frmMDIMainMenu.tbrLibrary.Buttons("New").Enabled = True
        frmMDIMainMenu.tbrLibrary.Buttons("Print").Enabled = True
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Enabled = True
        frmMDIMainMenu.mnuViewElip(0).Checked = False
        frmMDIMainMenu.mnuViewElip(3).Checked = True
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView0").Value = tbrUnpressed
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView1").Value = tbrUnpressed
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView2").Value = tbrUnpressed
        frmMDIMainMenu.tbrLibrary.Buttons("lvwView3").Value = tbrPressed
    End If
    GoTo End_Sub
Error_Handler:
DisplayError ""
Resume End_Sub
End_Sub:
    
End Sub


'===========================================================================
' DESC:     Center current form relative to Parent form, or if there are no
'           Parent form it will centered relative to screen
'===========================================================================
Public Sub CenterForm(frm As Form, Optional vParent As Variant)
  Dim oParent    As Object
  Dim iMode%
  Dim iLeft%
  Dim iTop%
  If IsMissing(vParent) Then
    Set oParent = Screen
  ElseIf TypeOf vParent Is Screen Or TypeOf vParent Is Form Then
    Set oParent = vParent
  Else
    Exit Sub
  End If
    If TypeOf oParent Is Form Then
      iLeft = oParent.Left
      iTop = oParent.Top
    End If
  frm.Move iLeft + (oParent.Width - frm.Width) / 2, (iTop + (oParent.Height - frm.Height) / 2) - 600
End Sub


'===========================================================================
' DESC:     Change the state of ToolBar and Menu of View.
' PARAMS:
'===========================================================================
Public Sub changeViewMain(bView As Variant)
    Dim ctrFrm As Form
  ' Change state of ToolBar and Menu of view
  For fiIndex% = 0 To 3
    If bView = "lvwView" & CStr(fiIndex%) Then
      frmMDIMainMenu.tbrLibrary.Buttons("lvwView" & CStr(fiIndex%)).Value = tbrPressed
      frmMDIMainMenu.mnuViewElip(fiIndex%).Checked = 1
    Else
      frmMDIMainMenu.tbrLibrary.Buttons("lvwView" & CStr(fiIndex%)).Value = tbrUnpressed
      frmMDIMainMenu.mnuViewElip(fiIndex%).Checked = 0
    End If
  Next fiIndex%
    
    ' Detect current Form to change the ListView.View
    Select Case gtMainLvw               ' Optimize to Funct GetCurrentForm
        Case "MainAcquisition"
            Set ctrFrm = frmMainAcquisition
        Case "MainCataloging"
            Set ctrFrm = frmMainCataloging
        Case "MainCirculation"
            Set ctrFrm = frmMainCirculation
        Case "MainPatron"
            Set ctrFrm = frmMainPatron
        Case "MainUtility"
            Set ctrFrm = frmMainUtility
        Case "RecPublisher"
            Set ctrFrm = frmPublisher
        Case "RecVendor"
            Set ctrFrm = frmVendor
        Case "RecSupply"
            Set ctrFrm = frmSupply
        Case "RecMaterial"
            Set ctrFrm = frmMaterial
        Case "RecMaterialType"
            Set ctrFrm = frmMaterialType
        Case "RecPlacement"
            Set ctrFrm = frmPlacement
        Case "RecSubject"
            Set ctrFrm = frmSubject
        Case "RecLanguage"
            Set ctrFrm = frmLanguage
'        Case "RecLoan"
'            Set ctrFrm = frmLoan
'        Case "RecReturn"
'            Set ctrFrm = frmReturn
        Case "RecPatron"
            Set ctrFrm = frmMember
            
        Case "RecMemberType"
            Set ctrFrm = frmMemberType
    End Select
  
  ' Change view of ListView
    If gtMainLvw Like "Main*" Then
        ctrFrm.lvwMain.View = CInt(Right(bView, 1))
    ElseIf gtMainLvw Like "Rec*" Then
        ctrFrm.lvwLibrary.View = CInt(Right(bView, 1))
    End If
  
End Sub


Public Function gtConvBoolString(ptBool As String) As String
    If ptBool Then
        gtConvBoolString = "Yes"
    Else
        gtConvBoolString = "No"
    End If
End Function

Public Sub ShowFirstGrp()
    frmContent.imgLogo.Visible = False
    frmContent.piclibrary.Visible = True
    frmContent.grplibrary.Visible = True
    frmMDIMainMenu.mnuGMat_Lang_In_Click
    frmContent.imgLogo.ZOrder 0
End Sub


'===========================================================================
'   DESC    :   Search through ListIndex Property of Object, with compare of
'               a search string
'   Apply To:   ComboBox, DirListBox, DriveListBox, FileListBox, ListBox
'===========================================================================
Public Sub srchLikeLstItem(pCtrFrm As Form, objLst As Control, ptSrch$)
    ptSrch$ = ptSrch$ & "*"
    If ptSrch$ = "" Then
        objLst.ListIndex = 0
        Exit Sub
    End If
    For fiIndex% = 0 To objLst.ListCount - 1
        If objLst.List(fiIndex%) Like ptSrch$ Then
            objLst.ListIndex = fiIndex%
            Exit Sub
        End If
    Next fiIndex%
    If objLst.ListCount <> 0 Then objLst.ListIndex = 0 'objLst.ListCount - 1 'Not Found
End Sub


Public Function mReplaceCharacter(ByVal vstrOrigCharacters, _
                           ByVal vstrReplaceCharacters, _
                           ByVal vstrString)
    Dim strResult
    Dim intIndex
    Dim intOldIndex
    strResult = ""
    'traverse string
    intIndex = 1
    Do
        'get next occurance of orig characters
        intOldIndex = intIndex
        intIndex = InStr(intIndex, vstrString, vstrOrigCharacters)
        If (intIndex > 0) Then
            '*************
            'match found
            '*************
            strResult = strResult & Mid(vstrString, intOldIndex, _
            intIndex - intOldIndex)
            strResult = strResult & vstrReplaceCharacters
            intIndex = intIndex + Len(vstrOrigCharacters)
        Else
            '*************
            'no match
            '*************
            'get rest of string
            strResult = strResult & Mid(vstrString, intOldIndex)
            Exit Do
        End If
    Loop
    mReplaceCharacter = strResult
End Function



Public Sub cSelectRowLvw(pCtrLvw As Control)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(pCtrLvw.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Or LVS_EX_FULLROWSELECT
    r = SendMessageLong(pCtrLvw.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Sub



'===========================================================================
'   DESC    :   Search through ListIndex Property of Object, with compare of
'               a search string. Return True if found
'   Apply To:   ComboBox, DirListBox, DriveListBox, FileListBox, ListBox
'===========================================================================
Public Function fSrchLstItem(pCtrFrm As Form, objLst As Control, ptSrch$)
    If ptSrch$ = "" Then
        fSrchLstItem = False
        Exit Function
    End If
    For fiIndex% = 0 To objLst.ListCount - 1
        If StrComp(ptSrch$, objLst.List(fiIndex%), 0) = 0 Then
        fSrchLstItem = True
            Exit Function
        End If
        fSrchLstItem = False
    Next fiIndex%
End Function


' Prepare the current form shown from record display to multiple criteria search form
Public Sub ctrFormSeacrhtoRec()
    Dim ctrFrm As Form
    Set ctrFrm = ctrGetCurrentForm
    ctrFrm.cmd3DBCtrl(1).Visible = True     ' Edit
    ctrFrm.cmd3DBCtrl(3).Visible = True     ' Cancel
    ctrFrm.cmd3DBCtrl(2).Enabled = False      ' Enter
End Sub

' Prepare the current form shown from multiple criteria search form to record display
Public Sub ctrFormRectoSeacrh()
    Dim ctrFrm As Form
    Set ctrFrm = ctrGetCurrentForm
    ctrFrm.cmd3DBCtrl(1).Visible = False    ' Edit
    ctrFrm.cmd3DBCtrl(3).Visible = False    ' Cancel
    ctrFrm.cmd3DBCtrl(2).Enabled = True      ' Enter
    ' Cleared TextFields
    FieldControlNew ctrFrm
    ' Initialize Control Tags
    iFldTotal% = gLibDB.TableDefs(gtTableActive).Fields.Count
    For Index = 1 To iFldTotal% - 1
        ctrFrm.Controls(Index).Tag = ""
    Next Index
End Sub



Public Function ctrGetCurrentForm() As Form
    ' Detect current Form to change the ListView.View
    Select Case gtMainLvw               ' Optimize to Funct GetCurrentForm
        Case "MainAcquisition"
            Set ctrGetCurrentForm = frmMainAcquisition
        Case "MainCataloging"
            Set ctrGetCurrentForm = frmMainCataloging
        Case "MainCirculation"
            Set ctrGetCurrentForm = frmMainCirculation
        Case "MainPatron"
            Set ctrGetCurrentForm = frmMainPatron
        Case "MainUtility"
            Set ctrGetCurrentForm = frmMainUtility
        Case "RecPublisher"
            Set ctrGetCurrentForm = frmPublisher
        Case "RecVendor"
            Set ctrGetCurrentForm = frmVendor
        Case "RecSupply"
            Set ctrGetCurrentForm = frmSupply
        Case "RecMaterial"
            Set ctrGetCurrentForm = frmMaterial
        Case "RecMaterialType"
            Set ctrGetCurrentForm = frmMaterialType
        Case "RecPlacement"
            Set ctrGetCurrentForm = frmPlacement
        Case "RecSubject"
            Set ctrGetCurrentForm = frmSubject
        Case "RecLanguage"
            Set ctrGetCurrentForm = frmLanguage
'        Case "RecLoan"
'            Set ctrGetCurrentForm = frmLoan
'        Case "RecReturn"
'            Set ctrGetCurrentForm = frmReturn
        Case "RecPatron"
            Set ctrGetCurrentForm = frmMember
        Case "RecMemberType"
            Set ctrGetCurrentForm = frmMemberType
        Case Else
            MsgBox "Internal Error", vbCritical
    End Select
End Function
