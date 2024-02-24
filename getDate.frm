VERSION 4.00
Begin VB.Form frmGetDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Date"
   ClientHeight    =   2865
   ClientLeft      =   2430
   ClientTop       =   2295
   ClientWidth     =   3720
   Height          =   3270
   Icon            =   "getDate.frx":0000
   Left            =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Top             =   1950
   Width           =   3840
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   300
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fraDate 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton optTimeType 
         Caption         =   "Specific"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optTimeType 
         Caption         =   "Duration"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox cboDateStart 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cboDateStart 
         Height          =   315
         Index           =   1
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboDateStart 
         Height          =   315
         Index           =   2
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboDateEnd 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cboDateEnd 
         Height          =   315
         Index           =   1
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cboDateEnd 
         Height          =   315
         Index           =   2
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Caption         =   "Date Start:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Caption         =   "Date End:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGetDate"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Base 1
Private fIsFirstLoad As Boolean

Private Sub cboDateEnd_Change(Index As Integer)
    On Local Error Resume Next
    Dim xDate As Date
    Dim tTempDate As String

'    gtDatePeriod = "#" & cboDateStart(1) & " " & cboDateStart(0) & ", " & cboDateStart(2) & "#"
    
    If Not (cboDateEnd(1) = "") And Not (cboDateEnd(0) = "") Then
        tTempDate = cboDateEnd(1) & " " & cboDateEnd(0) & ", " & cboDateEnd(2)
        xDate = CDate(tTempDate)
        lbldate(1) = xDate
    Else
        tTempDate = cboDateEnd(2)
        lbldate(1) = tTempDate
    End If
    
    Select Case Index
        Case 1, 2
            showDayEnd
    End Select

End Sub

Private Sub cboDateEnd_Click(Index As Integer)
cboDateEnd_Change (Index)
End Sub

Private Sub cboDateStart_Change(Index As Integer)
On Local Error Resume Next
Dim xDate As Date
Dim tTempDate As String
    
    If Not (cboDateStart(1) = "") And Not (cboDateStart(0) = "") Then
        tTempDate = cboDateStart(1) & " " & cboDateStart(0) & ", " & cboDateStart(2)
        xDate = CDate(tTempDate)
        lbldate(0) = xDate
    Else
        tTempDate = cboDateStart(2)
        lbldate(0) = Format(tTempDate, "yyyy,mm,dd")
    End If
    
    Select Case Index
        Case 1, 2
            showDayStart
    End Select
End Sub
Public Sub showDayStart()
If cboDateStart(2) = "" Then Exit Sub
    If (CInt(cboDateStart(2)) Mod 4) = 0 Then
        gDayMonth(2).iDay = 29
    Else
        gDayMonth(2).iDay = 28
    End If
    cboDateStart(0).Clear
    For iIndex = 1 To (gDayMonth(cboDateStart(1).ListIndex + 1).iDay)
        cboDateStart(0).AddItem iIndex
    Next iIndex
    cboDateStart(0).ListIndex = 0
End Sub

Public Sub showDayEnd()
If cboDateEnd(2) = "" Then Exit Sub
    If (CInt(cboDateEnd(2)) Mod 4) = 0 Then
        gDayMonth(2).iDay = 29
    Else
        gDayMonth(2).iDay = 28
    End If
    cboDateEnd(0).Clear
    For iIndex = 1 To (gDayMonth(cboDateEnd(1).ListIndex + 1).iDay)
        cboDateEnd(0).AddItem iIndex
    Next iIndex
    cboDateEnd(0).ListIndex = 0
End Sub

Private Sub cboDateStart_Click(Index As Integer)
    cboDateStart_Change (Index)
    
End Sub

Private Sub cmdSelect_Click()
    If optTimeType(0).Value = True Then
        gtDatePeriod = gtRepSearchItem & " = Date(" & Year(lbldate(0)) & "," & _
            Month(lbldate(0)) & "," & Day(lbldate(0)) & ")"
    ElseIf optTimeType(1).Value = True Then
        gtDatePeriod = gtRepSearchItem & " >= Date(" & Year(lbldate(0)) & "," & _
            Month(lbldate(0)) & "," & Day(lbldate(0)) & _
            ") AND " & gtRepSearchItem & " <=  Date(" & Year(lbldate(1)) & "," & _
            Month(lbldate(1)) & "," & Day(lbldate(1)) & ")"
    End If
    
    Init_Report
    Unload Me
End Sub

Private Sub Form_Load()
'    Me.Show
    On Local Error GoTo Error_Handler
    fIsFirstLoad = True
'    Dim iIndex As Integer
    Dim iIndex As Integer
    
    
'    InitDB getLibSet("Database")
    
    For iIndex = LBound(gDayMonth) To UBound(gDayMonth)
        cboDateStart(1).AddItem gDayMonth(iIndex).tMonth
        cboDateEnd(1).AddItem gDayMonth(iIndex).tMonth
    Next
    cboDateStart(1).ListIndex = 0
    cboDateEnd(1).ListIndex = 0
    
        
    For iIndex = (CInt(Year(gxYearStart)) - 3) To (CInt(Year(gxYearEnd)) + 3)
        cboDateStart(2).AddItem iIndex
        cboDateEnd(2).AddItem iIndex
        
    Next
    cboDateStart(2).ListIndex = 0
    cboDateEnd(2).ListIndex = cboDateEnd(2).ListCount - 1
    fIsFirstLoad = False
    GoTo End_Sub
Error_Handler:
    DisplayError ""
    Resume End_Sub
End_Sub:
End Sub

Private Sub optTimeType_Click(Index As Integer)
Dim iIndex As Integer
    Select Case Index
        Case 0
            cboDateEnd(0).Enabled = False
            cboDateEnd(1).Enabled = False
            cboDateEnd(2).Enabled = False
            lblInfo(0) = "Date:"
        Case 1
'            For iIndex = LBound(cboDateEnd) To UBound(cboDateEnd)
'                cboDateEnd(iIndex).Enabled = False
'            Next
            cboDateEnd(0).Enabled = True
            cboDateEnd(1).Enabled = True
            cboDateEnd(2).Enabled = True
            lblInfo(0) = "Date Start:"
    End Select
End Sub
