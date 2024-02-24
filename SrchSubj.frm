VERSION 4.00
Begin VB.Form frmSearchSubject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search by Subjects"
   ClientHeight    =   5775
   ClientLeft      =   2070
   ClientTop       =   1860
   ClientWidth     =   6840
   Height          =   6180
   Icon            =   "SrchSubj.frx":0000
   Left            =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Top             =   1515
   Width           =   6960
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3960
      Width           =   6615
   End
   Begin VB.ListBox lstIncludedFields 
      Height          =   2400
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox lstFields 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   6720
      X2              =   120
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Lbl 
      Caption         =   "Access No.    Title"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   6615
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
      i1              =   "SrchSubj.frx":000C
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
      i1              =   "SrchSubj.frx":0503
   End
   Begin Threed.SSCommand cmd3DFind 
      Height          =   300
      Left            =   4680
      TabIndex        =   7
      Top             =   3000
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Find..."
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
   Begin VB.Label lblInfo2 
      Caption         =   "Searched Subjects"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblInfo1 
      Caption         =   "A&vailable Subjects"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
   Begin Threed.SSCommand cmd3Close 
      Height          =   300
      Left            =   5760
      TabIndex        =   4
      Top             =   3000
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
   Begin Threed.SSCommand cmd3AddRemove 
      Height          =   300
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "<- &Remove"
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
      MouseIcon       =   "SrchSubj.frx":09FA
   End
   Begin Threed.SSCommand cmd3AddRemove 
      Height          =   300
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Add ->"
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
Attribute VB_Name = "frmSearchSubject"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmSubjectMat
'***************************************************************************


Private fDirty As Boolean       ' Flag to indicate changes


Private Sub cmd3AddRemove_Click(Index As Integer)
  Dim i As Integer
  Select Case Index
    Case 0
      If lstFields.ListIndex = -1 Then Exit Sub
      For i = lstFields.ListCount - 1 To 0 Step -1
        If lstFields.Selected(i) = True Then
          lstIncludedFields.AddItem lstFields.List(i)
'          lstFields.RemoveItem i
        End If
      Next
    Case 1
      If lstIncludedFields.ListIndex = -1 Then Exit Sub
      For i = lstIncludedFields.ListCount - 1 To 0 Step -1
        If lstIncludedFields.Selected(i) = True Then
 '         lstFields.AddItem lstIncludedFields.List(i)
          lstIncludedFields.RemoveItem i
        End If
      Next
End Select
End Sub

Private Sub cmd3Close_Click()
    Unload Me
End Sub


Private Function dbIsExist(ByVal Acq_Num As Single, ptSubject As String) As Boolean
    dbIsExist = False
    SQL = "SELECT SUBJECT_CODE FROM [SUBJECT-MATERIAL] WHERE ACQUISITION_NUM = '" & Acq_Num & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    If Not gLibSnap.EOF Then
        While Not gLibSnap.EOF
            If ptSubject = gLibSnap.Fields("SUBJECT_CODE") Then
                dbIsExist = True
                Exit Function
            End If
            gLibSnap.MoveNext
        Wend
    End If
End Function


Private Sub cmd3DFind_Click()
    On Error Resume Next
    Dim tSubject            As String
    Dim iIndex              As Integer
    Dim tFound              As String
    Dim libSnap             As Recordset
    Dim subLibSnap          As Recordset
    Dim tDummy              As String
    Dim tPrevios            As String
    
    'Initialization
    tSubject = ""
    list1.Clear
    DoEvents
    tFound = ""
    
    For iIndex = 0 To lstIncludedFields.ListCount - 1
        tSubject = tSubject & "'" & Parse(lstIncludedFields.List(iIndex), 1, ",") & "'" & ", "
    Next
    tSubject = Trim(tSubject)
    If lstIncludedFields.ListCount = 0 Then Exit Sub
    
    tSubject = Left(tSubject, Len(tSubject) - 1)
    
    SQL = "SELECT ACQUISITION_NUM FROM [SUBJECT-MATERIAL] WHERE SUBJECT_CODE IN (" & tSubject & ")"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    If Not gLibSnap.EOF Then
        While Not gLibSnap.EOF
            tFound = gLibSnap.Fields("ACQUISITION_NUM")
            SQL = "SELECT * FROM [SUBJECT-MATERIAL] WHERE ACQUISITION_NUM = '" & tFound & "' AND SUBJECT_CODE IN (" & tSubject & ")"
            Set libSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            If Not libSnap.EOF Then
                libSnap.MoveLast
                If libSnap.RecordCount = lstIncludedFields.ListCount Then
                    SQL = "SELECT TITLE FROM MATERIAL WHERE ACQUISITION_NUM = '" & tFound & "'"
                    Set subLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
                    tDummy = tFound & vbTab & vbTab & subLibSnap.Fields(0)
                    If tDummy <> tPrevious Then list1.AddItem tDummy
                    tPrevious = tDummy
                End If
            End If
             gLibSnap.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    fillLstAllSubject
End Sub


Private Sub fillLstAllSubject()
    Dim tSubject As String
    Dim fEqualString As Boolean
    lstFields.Clear
    
    SQL = "SELECT SUBJECT_CODE, SUBJECT_DESC FROM [SUBJECT]"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    With gLibSnap
        While Not .EOF
            tSubject = CStr(.Fields(0)) & ", " & CStr(.Fields(1))
            fEqualString = fSearchLstItem(frmMaterial.lstSubject, tSubject)
            If Not fEqualString Then
                lstFields.AddItem tSubject
            Else
                lstIncludedFields.AddItem tSubject
            End If
            .MoveNext
        Wend
    End With
End Sub


'===========================================================================
'   DESC    :   Search through List control of a given search string; return
'               True if found
'   Apply To:   ComboBox, DirListBox, DriveListBox, FileListBox, ListBox
'===========================================================================
Private Function fSearchLstItem(objLst As Control, ptSrch$) As Boolean
    For fiIndex% = 0 To objLst.ListCount - 1
        If StrComp(ptSrch$, objLst.List(fiIndex%), 0) = 0 Then
            fSearchLstItem = True
            Exit Function
        Else
            fSearchLstItem = False
        End If
    Next fiIndex%
End Function
