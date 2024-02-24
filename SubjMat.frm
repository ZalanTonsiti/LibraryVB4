VERSION 4.00
Begin VB.Form frmSubjectMat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subject"
   ClientHeight    =   3000
   ClientLeft      =   2220
   ClientTop       =   1740
   ClientWidth     =   8040
   Height          =   3405
   Icon            =   "SubjMat.frx":0000
   Left            =   2160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Top             =   1395
   Width           =   8160
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
   Begin VB.Label lblInfo2 
      Caption         =   "Material Subjec&ts"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   240
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
      MouseIcon       =   "SubjMat.frx":000C
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
Attribute VB_Name = "frmSubjectMat"
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
          lstFields.RemoveItem i
        End If
      Next
    Case 1
      If lstIncludedFields.ListIndex = -1 Then Exit Sub
      For i = lstIncludedFields.ListCount - 1 To 0 Step -1
        If lstIncludedFields.Selected(i) = True Then
          lstFields.AddItem lstIncludedFields.List(i)
          lstIncludedFields.RemoveItem i
        End If
      Next
End Select
End Sub

Private Sub cmd3Close_Click()
    updateSubjectMaterial
    updatefrmMaterial
    Unload Me
End Sub


Private Sub updatefrmMaterial()
    Dim iIndex As Integer
    frmMaterial.lstSubject.Clear
    For iIndex = 0 To lstIncludedFields.ListCount - 1
        frmMaterial.lstSubject.AddItem lstIncludedFields.List(iIndex)
    Next iIndex
End Sub


Private Sub updateSubjectMaterial()
    Dim iIndex As Integer
    Dim tSubject As String
    
    SQL = "SELECT * FROM [SUBJECT-MATERIAL] WHERE ACQUISITION_NUM = '" & gtAcquisition_Num & "'" '& _
        " AND SUBJECT_CODE = '" & tSubject & "'"
    Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
    If Not gLibDyna.EOF Then
        While Not gLibDyna.EOF
            gLibDyna.Delete
            gLibDyna.MoveNext
        Wend
    End If
    
    For iIndex = 0 To lstIncludedFields.ListCount - 1
        tSubject = Parse(CStr(lstIncludedFields.List(iIndex)), 1, ",")
        Debug.Print "tSubject", tSubject
        SQL = "SELECT * FROM [SUBJECT-MATERIAL]"
         Set gLibDyna = gLibDB.OpenRecordset(SQL, dbOpenDynaset)
        With gLibDyna
            .AddNew
            .Fields("SUBJECT_CODE") = tSubject
            .Fields("ACQUISITION_NUM") = gtAcquisition_Num
            .Update
        End With
    Next iIndex
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


Private Sub Form_Load()
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
