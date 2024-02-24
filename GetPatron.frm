VERSION 4.00
Begin VB.Form frmGetPatronNum 
   Caption         =   "Select Patron Number"
   ClientHeight    =   2865
   ClientLeft      =   3165
   ClientTop       =   3090
   ClientWidth     =   4680
   Height          =   3270
   Icon            =   "GetPatron.frx":0000
   Left            =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4680
   Top             =   2745
   Width           =   4800
   Begin VB.TextBox txtPatronNum 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.ListBox lstPatronNum 
      Height          =   1620
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin Threed.SSCommand cmd3Select 
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Select"
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
   Begin VB.Label lblPatronName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblInfoPatronName 
      Alignment       =   1  'Right Justify
      Caption         =   "Patron Name"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblInfoPatronNum 
      Alignment       =   1  'Right Justify
      Caption         =   "Patron Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmGetPatronNum"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' frmGetPatronNum
' Optimize : Could be used as common dialog to select Patron

Private Sub cmd3Select_Click()
On Local Error GoTo Error_Handler
    gtGetPatron_Num = txtPatronNum
    Unload Me
    GoTo End_Sub:
Error_Handler:
Resume End_Sub
End_Sub:
End Sub

Private Sub Form_Load()
    Select Case gtMainLvw
        Case "RecPatron"
            dbRetrieve lstPatronNum, "PATRON", "PATRON_NUM", 1
        Case "GenReport"
            dbRetrieve lstPatronNum, "LOAN", "PATRON_NUM", 1
    End Select
End Sub


Private Sub lstPatronNum_DblClick()
    txtPatronNum = lstPatronNum
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
        End If
    ElseIf Not (Len(txtPatronNum)) Then ' Blank entry
        lblPatronName = ""
    End If
    GoTo End_Sub
Error_Handler:
    DisplayError gtLibErr(0)
End_Sub:
End Sub
