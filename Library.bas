Attribute VB_Name = "mLibrary"

'===========================================================================
'   mLibrary
'   DESC:   Library Module
'===========================================================================

Public gLibWS                       As Workspace
Public gLibDB                       As Database
Public gLibDyna                     As Recordset
Public gLibSnap                     As Recordset
Public gLibTable                    As Recordset
Public gDbName                      As String
Public giTotalReport                As Integer
Public gtRepSearchItem              As String
Public gtFileNameAnnUpdt            As String

'Public gtCurrUsr$              ' Current User
'Public gtCurrTable As String    ' Current Correspond Table

Public gxYearStart As Date
Public gxYearEnd As Date
Public gCtrFrm As Form

Public gtTableActive                As String
Public gtTableIndex                 As String
Public gtTableIndex2                As String
Public gtCurrentIndex               As String
Public gfSearchMode                 As Boolean

Public gfNewStatus                  As Boolean
Public gfSQLCase                    As Boolean     ' SQL Flag 0 - String; 1 - Number
Public gfSort                       As Boolean
Public gtMainLvw                    As String
Public tSelectedItem                As String
Public Const giClaimDuration% = 7
Public Const giDateDefaultValue     As Date = "1/1/1980"
Public giAnnUpdate                  As Integer  ' Trace current step of AnnuAl Update
Public gtGetPatron_Num              As String
Public gtAcquisition_Num            As String
Public gtDatePeriod                 As String

' Application type definitions
Type defReport
    Name As String
    SelFormula As String
    Parameter As String
End Type

Type defModule
    tName As String
    tDescription As String
End Type

Public gReportFile() As defReport
Public gtLibErr() As String         ' Appilcation Generated Error

Public gtModule() As defModule

Type defDayMonth
    tMonth As String
    iDay As Integer
End Type

Public gDayMonth(1 To 12) As defDayMonth

'===========================================================================
'   DESC    :   SECURITY
'===========================================================================
Public strPasswd      As String
Public strUsrName     As String
Public currUsr        As String           ' Current User
Public gfLogConfirm   As Boolean    ' True - Second Confirmation
                                    ' False - Normal System Log

' System Settings
Public boolSetPwdLen As Boolean
Public intSetPwdLen As Integer '2B
Public boolSetPwdAlphaNum As Boolean

Public Const strSystem = "Library Information System" '"LIBRARY INFORMATION SYSTEM"
Public Const strExecName = "ILIS - by AI NM"
Public m_hStatBar As Long

' Notes :
' Please refer to API Viewer (C:\Program Files\Microsoft Visual Basic\winapi\apilod32.exe)

' Win32 API declaration to get the current user of Windows
Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As _
  String, ByVal lpUserName As String, lpnLength As Long) As Long

' Win32 API to launch file with associated application. E.g (.htm|IE, .txt|Notepad) etc
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd _
  As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters _
  As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
  As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long

' NEXT VERSION
'   Error Add User - Cannot add Name$. A user with the name you specified already exists.
'   Specify a different username. vbCritical.

Declare Function CreateStatusWindow Lib "comctl32.dll" Alias _
  "CreateStatusWindowA" (ByVal style As Long, ByVal lpszText As String, ByVal _
  hWndParent As Long, ByVal wID As Long) As Long


'===========================================================================
' DESC:     Module for handling first run of the App
'===========================================================================
Public Sub FirstRun()
    MsgBox "This is your first time running " & gtEXEName & _
        ". A database will created for you now.", vbInformation, strSystem
    With frmLibOptions.cmlLibrary
        .Filter = "Library Database Files (*.mdb)|*.mdb"
        .FilterIndex = 0
        .Flags = FileOpenConstants.cdlOFNHideReadOnly
        .ShowSave
        If .FileName = "*.mdb" Or .FileName = "" Then Exit Sub
            ' File already exists, so ask if the user wants to overwrite the file.
            If Dir(.FileName) <> "" Then
                If MsgBox("Overwrite existing file?", vbYesNo + vbQuestion + _
                    vbDefaultButton2) = vbNo Then Exit Sub
                Kill .FileName
            End If
        CreateLibDB .FileName
    End With
    DoEvents
'    MsgBox "Please restart " & strSystem, vbInformation
'    End
End Sub


'===========================================================================
' DESC:     Module for initializing the database
'===========================================================================
Sub InitDB(FileName As String)
    Dim dbName As String
    On Local Error GoTo Error_Handler
    
    If Len(FileName) = 0 Then gDbName$ = getLibSet("Database") Else gDbName$ = FileName
    
    If GetSetting(strExecName, "Settings", "FirstRun") = "True" Then
        FirstRun
        DeleteSetting strExecName, "Settings", "FirstRun"
    End If

    If gDbName$ = "0" Or gDbName$ = "" Or Dir(gDbName$) = "" Then
            MsgBox "There are no default database specify.", vbInformation, strSystem
            
            With frmLibOptions.cmlLibrary
                .Filter = "Library Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
                .FilterIndex = 0
                .Flags = FileOpenConstants.cdlOFNHideReadOnly
                .ShowOpen
            If .FileName = "*.mdb" Then
                Exit Sub
            ElseIf .FileName = "" Then
                MsgBox "You must specify a default Database. " & _
                    "This application will END now.", vbInformation, strSystem
                End
            End If
                
                gDbName$ = .FileName
                If MsgBox("Do you want to specify the database " & .FileName & _
                    " to default database.", vbYesNo + vbQuestion + _
                    vbDefaultButton1, strSystem) = vbYes Then _
                    SaveSetting strExecName, "Settings", "Database", .FileName
        
            End With
    End If
    
    Set gLibWS = DBEngine.CreateWorkspace("Workspace", "Admin", "")
    
    '   Open DataBase with Exclusive and Read/Write access
    Set gLibDB = gLibWS.OpenDatabase(gDbName$, True, False)
    
    If Not fCheckValidDB Then
        MsgBox "The database specified is not compliant with this application.", vbcritcal, strSystem
        End
    End If
    
    GoTo End_Sub
Error_Handler:
    DisplayError gtLibErr(0)
    Resume End_Sub
End_Sub:
End Sub

' DESC:     Retrieve Data from Table and display to ComboBox, ListBox

Public Sub Init_Report()
    ReDim gReportFile(1 To giTotalReport) As defReport
    
    gReportFile(1).Name = "List of Members"
    gReportFile(1).SelFormula = ""
    gReportFile(1).Parameter = 1

    gReportFile(2).Name = "Member Cards"
    gReportFile(2).SelFormula = ""
    gReportFile(2).Parameter = 1

    gReportFile(3).Name = "List of Vendors"
    gReportFile(3).SelFormula = ""
    gReportFile(3).Parameter = 1

    gReportFile(4).Name = "List of Material Type"
    gReportFile(4).SelFormula = ""
    gReportFile(4).Parameter = 1

    gReportFile(5).Name = "List of Languages"
    gReportFile(5).SelFormula = ""
    gReportFile(5).Parameter = 1

    gReportFile(6).Name = "List of Subject"
    gReportFile(6).SelFormula = ""
    gReportFile(6).Parameter = 1

    gReportFile(7).Name = "List of Lost Material"
    gReportFile(7).SelFormula = "{MATERIAL.EXISTENCE} = FALSE"
    gReportFile(7).Parameter = 1

    gReportFile(8).Name = "Booking History"
    gReportFile(8).SelFormula = ""
    gReportFile(8).Parameter = 1

    gReportFile(9).Name = "Amount of fine"
    gReportFile(9).SelFormula = "{LOAN.TOT_FINE} > 0"
    gReportFile(9).Parameter = 1

    gReportFile(10).Name = "Amount of fine collected based on date"
    gReportFile(10).SelFormula = gtDatePeriod
    gReportFile(10).Parameter = 3

    gReportFile(11).Name = "List of purchased materials based on Date"
    gReportFile(11).SelFormula = gtDatePeriod
    gReportFile(11).Parameter = 3

    gReportFile(12).Name = "Materials used by certain individuals"
    gReportFile(12).SelFormula = "{Material Used by certain Patron.PATRON_NUM} = '" & gtGetPatron_Num & "'"
    gReportFile(12).Parameter = 2
    
    ' select count(*) as [Material Used]
'    gReportFile(13).Name = "Materials used based on material type (borrowed)"
'    gReportFile(13).SelFormula = ""
'    gReportFile(13).Parameter = 1

'    gReportFile(14).Name = "Analysis of books stored in the library based on subjects"
'    gReportFile(14).SelFormula = ""
'    gReportFile(14).Parameter = 1

'    gReportFile(15).Name = "List of Overdue Material"
'    gReportFile(15).SelFormula = ""
'    gReportFile(15).Parameter = 1

End Sub



Public Sub Init_App_Error()
    ReDim gtLibErr(0 To 50)
    gtLibErr(0) = "Fatal Error. Application will be shutdown now."
    gtLibErr(4) = "Please set a default printer on this computer system."
    gtLibErr(12) = "There are no items to shown in this view"
    gtLibErr(25) = "Please recheck association these particular record with " & _
        "other modules. The complete elimination of the relationship will enable deletion."
    
End Sub

'===========================================================================
' Startup procedure
'   Check the minimum 800 X 600 pixels SCREEN resolution
'   Init DataBase
'   Init Report coleection
'   Show First Form
'===========================================================================
Sub Main()
On Local Error GoTo Error_Handler:
    Init_App_Error

    ' Check current Screen resolution
    If Not fCheckScreen Then
        MsgBox "This application design to work in 800 X 600 pixels environment.", vbInformation
        End
    End If
    
    ' Check if there is a default printer
    chkPrinter

    ' Detect first run
    If getLibSet("Copyright") <> "AI NM" Then
        'CheckReg Fail - Add Copyright Mark
        saveLibSet "Copyright", "AI NM"
        ' Save to registry indicate new install
        saveLibSet "FirstRun", "True"
        saveLibSet "Check Date", "0"
        saveLibSet "Last Access Date", Date
        saveLibSet "SelectRow", "0"
        saveLibSet "ShowHeader", "False"
    End If
    
    If Not CBool(getLibSet("Check Date")) Then
      If chkDateSecurity Then
        MsgBox "Last Access Date: " & CStr(getLibSet("Last Access Date")) _
            & vbCrLf & "Current System Date: " & CStr(Date) & vbCrLf & _
            CDate(getLibSet("Last Access Date")) - Date
      End If
    End If
    
    gDbName$ = (getLibSet("Database"))
    
    giTotalReport = 13
    
    InitDB gDbName$
    DoEvents
    gfLogConfirm = False    ' Init Second confirmation of Admin Password
    frmAbout.Show
    DoEvents
    Init_Report
    ' Get current time
    xStart = Now
    ' Wait for five seconds before proceed
    While DateDiff("s", xStart, Now) <= 5
    DoEvents
        If DateDiff("s", xStart, Now) < 3 Then
            frmAbout.lblInfo(0) = "LIBRARY INFORMATION SYSTEM" & vbCrLf & "Initializing application..."
        ElseIf DateDiff("s", xStart, Now) < 5 Then
            frmAbout.lblInfo(0) = "LIBRARY INFORMATION SYSTEM" & vbCrLf & "Refreshing desktop..."
        End If
    
    Wend
    gfSearchMode = False
    ChDir App.Path
    Unload frmAbout
    frmMDIMainMenu.Show
    GoTo End_Sub
Error_Handler:
    DisplayError ""
End_Sub:
End Sub


'===========================================================================
' DESC:     Display error message with custom message, Error Number and
'           Description
'===========================================================================
Public Sub DisplayError(gError$)
    Select Case Err.Number
        Case 524
            gError$ = gError$ & " - Related entry does not exist"
        Case 3201
            gError$ = gError$ & " - Please check validity of your entry"
        Case 3315
            gError$ = gError$ & " - Please enter valid entry"
        Case 3421
            gError$ = gError$ & " - Please check the Data Type of the " & _
                "entries. You may have left an entry blank"
        Case 3356
            gError$ = gError$ & " - Please close all other application."
        Case 3022
            gError$ = gError$ & " - Duplicated record. Unable to Update."
            Select Case gtTableActive
                Case "LOAN"
                    gError$ = gError$ & vbCrLf & " The paton had borrow " & _
                        "and return this item in the same day."
                End Select
    End Select
    MsgBox gError$ & vbCrLf & vbCrLf & "VB" & Err.Number _
        & ": " & Err.Description, vbExclamation
    Err.Clear
End Sub


'===========================================================================
' DESC:     Retrieve Data from Table and display to ComboBox, ListBox
'===========================================================================
Public Sub dbRetrieve(ctrl As Control, tTable$, tField$, Optional vOption As Variant)
    ctrl.Clear
    
    If IsMissing(vOption) Then
        ctrl.AddItem ""
    End If
    
    SQL = "SELECT DISTINCT [" & tTable$ & "].[" & tField & "] FROM [" & tTable & "]"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    While Not gLibSnap.EOF
        If Not IsNull(gLibSnap.Fields(tField)) Then
            ctrl.AddItem gLibSnap.Fields(tField)
        End If
        gLibSnap.MoveNext
    Wend
    If ctrl.ListCount Then ctrl.ListIndex = 0

End Sub


'===========================================================================
' DESC:     Check screen resolution if 800 X 600
'===========================================================================
Public Function fCheckScreen() As Boolean
    iHeight = Screen.Height / Screen.TwipsPerPixelY
    iWidth = Screen.Width / Screen.TwipsPerPixelX
If iHeight < 600 Or iWidth < 800 Then
    fCheckScreen = False
Else
    fCheckScreen = True
End If
    
End Function


'===========================================================================
'   FUNCTION    :   Parse
'   DESC        :   Parsing a substring from a string with a specified
'                   Delimiter
'   Note        :   High Cyclomatic
'===========================================================================
Public Function Parse(ByVal ptWhole$, piPos%, DELIMITER$, _
    Optional pfTrimSource As Variant, Optional pfTrimSink As Variant) As String
    
    
On Error Resume Next
    Dim iLenWhole%
    Dim iPlace%
    Dim tCurrentChar$
    Dim iCounter1%
    Dim iCounter2%
    Dim iStart%
    Dim iStop%
    If IsMissing(pfTrimSource) Then
        ptWhole$ = DELIMITER + ptWhole$ + DELIMITER
    ElseIf pfTrimSource Then
        ptWhole$ = DELIMITER + Trim(ptWhole$) + DELIMITER
    End If
    iLenWhole% = Len(ptWhole$)
    iPlace% = 0
    For iCounter1% = 1 To iLenWhole%
        tCurrentChar$ = Mid$(ptWhole$, iCounter1%, 1)
        If tCurrentChar$ = DELIMITER Then iPlace% = iPlace% + 1
        If iPlace% = piPos% Then
            iStart% = iCounter1% + 1
            Exit For
        End If
    Next
    For iCounter2% = iStart% To iLenWhole%
        tCurrentChar$ = Mid$(ptWhole$, iCounter2%, 1)
        If tCurrentChar$ = DELIMITER Then iPlace% = iPlace% + 1
        If iPlace% = piPos% + 1 Then
            iStop% = iCounter2% - iStart%
            Exit For
        End If
    Next
    If iStop% = 0 Then
        If IsMissing(pfTrimSink) Then
            Parse = Mid$(ptWhole$, iStart%)
        ElseIf pfTrimSink Then
            Parse = Trim(Mid$(ptWhole$, iStart%))
        End If
          If InStr(Parse, DELIMITER) Then Parse = ""
    Else
        If IsMissing(pfTrimSink) Then
            Parse = Mid$(ptWhole$, iStart%, iStop%)
        ElseIf pfTrimSink Then
            Parse = Trim(Mid$(ptWhole$, iStart%, iStop%))
        End If
        If InStr(Parse, DELIMITER) Then Parse = ""
    End If
  On Error GoTo 0
  On Error Resume Next
End Function


'===========================================================================
' DESC:     EncryptPassword
'           Function to get current UserID of Windows 95
' PARAMS:
'===========================================================================
Public Function UserID() As String
  Dim sUserNameBuffer As String * 255
  sUserNameBuffer = Space(255)
  Call WNetGetUser(vbNullString, sUserNameBuffer, 255&)
  UserID = Left$(sUserNameBuffer, InStr(sUserNameBuffer, vbNullChar) - 1)
End Function


'===========================================================================
' DESC:     EncryptPassword
'           Encryption function
' PARAMS:   Number                            DecryptedPassword
' SOURCE:   VB Planet
'===========================================================================
Public Function EncryptPassword(Number As Byte, DecryptedPassword As String)
  Dim Password As String, Counter As Byte
  Dim Temp As Integer
  ' The number passed may be 0 - 60 (seconds) mod by 10 to get range 0 - 10
  Number = Number Mod 10
  Counter = 1         ' See also : Option Base 1
  Do Until Counter = Len(DecryptedPassword) + 1
    Temp = Asc(Mid(DecryptedPassword, Counter, 1))
    If Counter Mod 2 = 0 Then
      Temp = Temp - Number
    Else
      Temp = Temp + Number
    End If
    Temp = Temp Xor (10 - Number)   ' Swap
    Password = Password & Chr$(Temp)
    Counter = Counter + 1
  Loop
  EncryptPassword = Password
End Function


'===========================================================================
' DESC:     DecryptPassword
'           Decryption function
' PARAMS:   Number                            EncryptedPassword
' SOURCE:   VB Planet
'===========================================================================
Function DecryptPassword(Number As Byte, EncryptedPassword As String)
  Dim Password As String, Counter As Byte
  Dim Temp As Integer
  ' The number passed may be 0 - 60 (seconds) mod by 10 to get range 0 - 10
  Number = Number Mod 10
  Counter = 1         ' See also : Option Base 1
  Do Until Counter = Len(EncryptedPassword) + 1
    Temp = Asc(Mid(EncryptedPassword, Counter, 1)) Xor (10 - Number) 'Swap
    If Counter Mod 2 = 0 Then
      Temp = Temp + Number
    Else
      Temp = Temp - Number
    End If
    Password = Password & Chr$(Temp)
    Counter = Counter + 1
  Loop
  DecryptPassword = Password
End Function


'===========================================================================
' DESC:     getCurrUsrSecDate
'           Function to extract Second out of user's creation date
'           Will generate error if the user logged is not listed (hacked?)
' PARAMS:
'===========================================================================
Public Function getCurrUsrSecDate() As Integer
  On Local Error GoTo Err_Handler
  Dim CurrUsrDate As Variant
  CurrUsrDate = GetSetting(strExecName, "Created", frmPwd.txtUsr.Text)
  getCurrUsrSecDate = Second(CurrUsrDate)
  Exit Function
Err_Handler:
  loginState = False
End Function


'===========================================================================
' DESC:     chkIsAlpha
'           Function to detect whether alphabet contain in string
' PARAMS:   strToDo       - String
'===========================================================================
Public Function chkIsAlpha(strToDo As String) As Boolean
  Dim Index As Integer
  chkIsAlpha = False
  For Index = 1 To Len(strToDo)
    Select Case Asc(Mid$(strToDo, Index, 1))
      Case 65 To 90, 97 To 122
        chkIsAlpha = True
        Exit Function
    End Select
  Next Index
'AZ 65-90 az 97-122 09 48-57
End Function


'===========================================================================
' DESC:     chkIsNum
'           Function to detect whether numeric contain in string
' PARAMS:   strToDo       - String
'===========================================================================
Public Function chkIsNum(strToDo As String) As Boolean
  Dim Index As Integer
  chkIsNum = False
  For Index = 1 To Len(strToDo)
    Select Case Asc(Mid$(strToDo, Index, 1))
      Case 48 To 57
        chkIsNum = True
        Exit Function
    End Select
  Next Index
'AZ 65-90 az 97-122 09 48-57
End Function


'===========================================================================
' DESC:     chkIsAlphaNum
'           Function to detect whether alphabet and numeric contain in string
' PARAMS:   strToDo       - String
'===========================================================================
Public Function chkIsAlphaNum(strToDo As String) As Boolean
  If chkIsAlpha(strToDo) And chkIsNum(strToDo) Then
    chkIsAlphaNum = True
  Else
    chkIsAlphaNum = False
  End If
End Function


'===========================================================================
' DESC:     GetDirectoryName
' PARAMS:   pPath       - Path
'===========================================================================
Function GetDirectoryName(ByVal pPath As String) As String
    Dim Path As String
    Dim FoundPath As Boolean
    Dim Count As Integer
    On Local Error GoTo ErrDirectory
    ' Test path until the valid one is found
    FoundPath = False
    While Not FoundPath
        Path = Dir(pPath)
        If Not IsNull(Path) Then
            FoundPath = True
        Else
            GoSub MovePortion
        End If
        If pPath = "" Then
            Path = ""
            FoundPath = True
        End If
    Wend
    ' Return
    GetDirectoryName = pPath
    Exit Function
' Removing last portion
MovePortion:
    For Count = Len(pPath) To 1 Step -1
        If Mid(pPath, Count, 1) = "\" Then
            pPath = Left(pPath, Count - 1)
            Exit For
        End If
    Next Count
    Return
' Error handler
ErrDirectory:
    ' Return
    GetDirectoryName = Null
    Exit Function
End Function


' Assign the current path E.g. "C:\" or "C:\Test1"
Public Function GetPath() As String
  If Right(App.Path, 1) = "\" Then
    GetPath = App.Path
  Else
    GetPath = App.Path & "\"
  End If
End Function


Public Sub dbAnnualUpdate()
    On Local Error GoTo CloseError
    dbAnnUpdateStepOne
    GoTo End_Sub
CloseError:
    MsgBox "Error occurred while trying to close file, please retry.", vbCritical, strSystem
Resume End_Sub
End_Sub:
End Sub


Public Sub dbAnnUpdateStepOne()
    giAnnUpdate = 0
    With frmProgress
        .Caption = "STEP 1 OF 3"
        .lblInfo.Caption = "This step will save the content of the cleared transaction to " & _
            "files. Click Browse choose the file to be saved."
        .cmd3Progress(3).Enabled = True
        .cmd3Progress(3).Visible = True
        .Show
    End With
End Sub


Public Sub dbAnnupdateStepTwo()
    SaveFileAs
    giAnnUpdate = 2
    With frmProgress
        .Caption = "STEP 2 OF 3"
        .lblInfo.Caption = "This step will print the content of the cleared transaction to " & _
            "printer. Click Next to Preview. Click the Printer icon on the form to print the report."
        .cmd3Progress(3).Visible = False
        .txtinfo.Visible = False
        .Show
    End With
End Sub


Public Sub dbAnnupdateStepThree()
    giAnnUpdate = 3
    With frmProgress
        .Caption = "STEP 3 OF 3"
        .lblInfo.Caption = "This step will delete all cleared fine transactions in the Loan module"
        .Show
    End With
End Sub


Public Sub getFilename()
    giAnnUpdate = 1
    With frmLibOptions.cmlLibrary
        .Filter = "Text Documents (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 0
        .Flags = FileOpenConstants.cdlOFNHideReadOnly
        .ShowSave
        If .FileName = "*.txt" Or .FileName = "" Then Exit Sub
        ' File already exists, so ask if the user wants to overwrite the file.
        If Dir(.FileName) <> "" Then
            If MsgBox("Overwrite existing file?", vbYesNo + vbQuestion + _
                vbDefaultButton2) = vbNo Then
                Exit Sub
            Else
            frmProgress.txtinfo = .FileName
            Kill .FileName
            End If
        Else
            frmProgress.txtinfo = .FileName
        End If
    End With
    
End Sub
    


Public Sub SaveFileAs()
    On Local Error GoTo Error_Handler
    
    Open gtFileNameAnnUpdt For Output As #1     ' Open file for output.
    Dim ftfldvalue As String
    
    SQL = "SELECT * FROM LOAN ORDER BY DATE_RETURN DESC"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot) '
    
    giFldTotal% = gLibDB.TableDefs("LOAN").Fields.Count - 1
    
    For fiIndex% = 0 To giFldTotal% '- 1
        ftFldName$ = mReplaceCharacter("_", " ", CStr(gLibDB.TableDefs("LOAN").Fields(fiIndex%).Name))
        Write #1, ftFldName$; '& Chr$(44);
    
    Next fiIndex%
    
    Write #1,
    
    While Not gLibSnap.EOF
        
        For fiIndex% = 0 To giFldTotal%
            ' Filter settled  transcation only
            If IsNull(gLibSnap.Fields("DATE_RETURN")) Then Exit For
            
            ftFldName$ = CStr(gLibDB.TableDefs("LOAN").Fields(fiIndex%).Name)
            If Not IsNull(gLibSnap.Fields(ftFldName$)) Then
                ftfldvalue$ = gLibSnap.Fields(ftFldName$)
                Write #1, ftfldvalue$; '& Chr$(44);
            Else
                Write #1, 'Chr$(44);
            End If
        Next fiIndex%
        
        Write #1,
        
        gLibSnap.MoveNext
    Wend
Close #1    ' Close file.

GoTo End_Sub

Error_Handler:
    DisplayError ""
Resume End_Sub
End_Sub:
End Sub


'===========================================================================
' DESC:         Init the day of month in a year
'===========================================================================
Public Sub InitDayMonth()
    gDayMonth(1).tMonth = "January"
    gDayMonth(1).iDay = 31
    gDayMonth(2).tMonth = "February"
    gDayMonth(2).iDay = 29
    gDayMonth(3).tMonth = "March"
    gDayMonth(3).iDay = 31
    gDayMonth(4).tMonth = "April"
    gDayMonth(4).iDay = 30
    gDayMonth(5).tMonth = "May"
    gDayMonth(5).iDay = 31
    gDayMonth(6).tMonth = "Jun"
    gDayMonth(6).iDay = 30
    gDayMonth(7).tMonth = "July"
    gDayMonth(7).iDay = 31
    gDayMonth(8).tMonth = "August"
    gDayMonth(8).iDay = 31
    gDayMonth(9).tMonth = "September"
    gDayMonth(9).iDay = 30
    gDayMonth(10).tMonth = "October"
    gDayMonth(10).iDay = 31
    gDayMonth(11).tMonth = "November"
    gDayMonth(11).iDay = 30
    gDayMonth(12).tMonth = "December"
    gDayMonth(12).iDay = 31
End Sub

'===========================================================================
' DESC:         Create new database with module DBCreate
'===========================================================================
Public Sub CreateNewDB()

On Local Error GoTo Error_Handler
    With frmMDIMainMenu.cmlLibrary
        .Filter = "Library Database Files (*.mdb)|*.mdb"
        .FilterIndex = 0
        .Flags = FileOpenConstants.cdlOFNHideReadOnly
        .ShowSave
        
        If .FileName = "*.mdb" Then Exit Sub
            ' File already exists, so ask if the user wants to overwrite the file.
            If Dir(.FileName) <> "" Then
                If MsgBox("Overwrite existing file?", vbYesNo + vbQuestion + _
                    vbDefaultButton2) = vbNo Then Exit Sub
                Kill .FileName
            End If
        CreateLibDB .FileName
    
    End With
    MsgBox "A new database - " & frmMDIMainMenu.cmlLibrary.FileName & " successfully created.", vbInformation, strSystem
    GoTo End_Sub
Error_Handler:
    DisplayError "Error occurred while trying to close file, please retry."
    Resume End_Sub
End_Sub:
End Sub


'===========================================================================
' DESC:         Retrieve and check the "Settings" in Registry. Return 0
'               if not exist
' PARAMS:       ptSetName$
'===========================================================================
Public Function getLibSet(ptSetName$) As String
  getLibSet = GetSetting(strExecName, "Settings", ptSetName$, "0")
End Function


'===========================================================================
' DESC:         Save the "Settings" in Registry
' PARAMS:       ptKey$, ptSetting$
'===========================================================================
Public Sub saveLibSet(ptKey$, ptSetting$)
  SaveSetting strExecName, "Settings", ptKey$, ptSetting$
End Sub


'===========================================================================
' DESC:         Check the Existence of record for reports
'===========================================================================
Public Function CheckReportDate() As Boolean
    CheckReportDate = False
    Select Case gtTableActive
        Case "LOAN"
            SQL = "SELECT DATE_RETURN FROM LOAN WHERE TOT_FINE > 0 ORDER BY DATE_RETURN ASC"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            If gLibSnap.EOF Then
                MsgBox "No current record of item returned late.", vbInformation, strSystem
            ElseIf IsNull(gLibSnap.Fields("DATE_RETURN")) Then
                MsgBox "No current record of item returned.", vbInformation, strSystem
            ElseIf Not gLibSnap.EOF Then
                gLibSnap.MoveFirst
                gxYearStart = gLibSnap.Fields(0)
                gLibSnap.MoveLast
                gxYearEnd = gLibSnap.Fields(0)
                CheckReportDate = True
            End If
        Case "SUPPLY"
            SQL = "SELECT DATE_OF_SUPPLY FROM SUPPLY ORDER BY DATE_OF_SUPPLY ASC"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
            If IsNull(gLibSnap.Fields("DATE_OF_SUPPLY")) Then
                MsgBox "There are no record to be displayed"
            ElseIf Not gLibSnap.EOF Then
                gLibSnap.MoveFirst
                gxYearStart = gLibSnap.Fields(0)
                gLibSnap.MoveLast
                gxYearEnd = gLibSnap.Fields(0)
                CheckReportDate = True
            End If
        Case Else
            MsgBox "Internal Error"
    End Select

End Function


'===========================================================================
' DESC:         Check the validity of Library Database
'                The database specified checked thoroughly for each Table
'                Name, Field Name, Field Type; and Field Size if the Field
'                Type is Text. Return False if do not match with supplied
'                String Constant
'===========================================================================
Public Function fCheckValidDB() As Boolean
    Dim tDbID As String
    Dim tTableName As String
    Dim tTableID As String
    
    Const tDBOrgID As String = "LANGUAGELOANMATERIALMATERIAL-TYPEMEMBER-TYPEPATRONPLACEMENTPUBLISHERRESERVESUBJECTSUBJECT-MATERIALSUPPLYVENDORLANGUAGE_CODELANGUAGE_DESCPATRON_NUMACQUISITION_NUMDATE_BORROWDATE_RETURNTOT_FINECLEARED_FINEDATE_CLEARED_FINEACQUISITION_NUMMATERIAL_CODELANGUAGE_CODESHELF_CODEVENDOR_CODEPUBLISHER_CODECLASSIFICATION_NUMAUTHOR1AUTHOR2AUTHOR3ISBN_NUMEDITIONSINOPSISYEAR_PUBLISHEDPLACE_PUBLISHEDON-LOANRESERVE-LISTEXISTENCEREFERENCE_ONLYTITLEMATERIAL_CODEMATERIAL_DESCDURATION_OF_BORROWMEMBER_CODEMEMBER_DESCMAX_BORROWINGBORROWING_DURATIONFINE_RATE_PERDAYPATRON_NUMMEMBER_CODEIC_NUMNAMEADDRESSTELEPHONE_NUMTOT_ITEM_BORROWEDACTIVE_STATUSCOURSE_DEPTSHELF_CODESHELF_DESCPUBLISHER_CODEPUBLISHER_NAMEACQUISITION_NUMPATRON_NUMBOOKING_DATERECEIVE_DATEEXPIRY_BOOKING_DATESUBJECT_CODESUBJECT_DESCSUBJECT_CODEACQUISITION_NUMACQUISITION_NUMVENDOR_CODEDATE_OF_SUPPLYCOSTVENDOR_CODEVENDOR_NAMEADDRESSFAX_NUMBERTELEPHONE_OFFTELEPHONE_HPCONTACT_PERSON"
    
    tTableName = ""
    tTableID = ""
    fCheckValidDB = False

    For Each tbldef In gLibDB.TableDefs
        If InStr(tbldef.Name, "MSys") = 0 Then
        tTableName = tTableName & CStr(tbldef.Name)
        End If
    Next
    
    
    For Each tbldef In gLibDB.TableDefs
        If InStr(tbldef.Name, "MSys") = 0 Then
            For Each fld In tbldef.Fields
                tTableID = tTableID & CStr(fld.Name) '& fld.Type
            Next
        End If
    Next

    tDbID = tTableName + tTableID

    If tDBOrgID = tDbID Then fCheckValidDB = True

End Function


'===========================================================================
' DESC:         Check the existence of System's default printer
'===========================================================================
Public Sub chkPrinter()
    On Local Error GoTo Error_Handler
    Dim tDummy$
    tDummy = Printer.DeviceName
    GoTo End_Sub
Error_Handler:
    DisplayError gtLibErr(4)
    Resume End_Sub
End_Sub:
End Sub


'===========================================================================
' DESC:         Check the Current Date and Last Access Date
'===========================================================================
Public Function chkDateSecurity() As Boolean
    chkDateSecurity = True
    
    If CInt(CDate(getLibSet("Last Access Date")) - Date) >= 0 Then _
        chkDateSecurity = False
End Function
