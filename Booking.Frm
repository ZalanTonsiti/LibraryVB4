VERSION 4.00
Begin VB.Form frmBooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking"
   ClientHeight    =   2655
   ClientLeft      =   975
   ClientTop       =   4245
   ClientWidth     =   10110
   Height          =   3060
   Icon            =   "Booking.frx":0000
   Left            =   915
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   10110
   Top             =   3900
   Width           =   10230
   Begin VB.Frame fraBooking 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   1
         Left            =   6600
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.ListBox lstAcqNum 
         Height          =   840
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.ListBox lstPatronNum 
         Height          =   840
         Left            =   6600
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Acquisition Number"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Patron Number"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Material Title"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin Threed.SSCommand cmd3Booking 
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Submit &Booking"
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
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   6600
         TabIndex        =   6
         Top             =   1560
         Width           =   3135
      End
      Begin Threed.SSCommand cmd3Cancel 
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
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
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'***************************************************************************
' frmBooking
'***************************************************************************


Private Sub cmd3Booking_Click()
On Local Error GoTo Error_Handler

    ' Check ON-LOAN status of a material
    SQL = "SELECT MATERIAL.[ON-LOAN] FROM MATERIAL WHERE " _
        & "ACQUISITION_NUM = '" & txtinfo(0) & "'"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    
    If Not (gLibSnap.Fields("ON-LOAN")) Then
        MsgBox "The material to be booked is not on loan", _
            vbInformation
        txtinfo(0) = ""
    GoTo End_Sub
    End If

    If Len(lblInfo(0)) Then
        ' Update RESERVE table
        Set gLibDyna = gLibDB.OpenRecordset("RESERVE", dbOpenDynaset)
        With gLibDyna
            .AddNew
                .Fields("ACQUISITION_NUM") = txtinfo(0)
                .Fields("PATRON_NUM") = txtinfo(1)
                .Fields("BOOKING_DATE") = Date
            .Update
        End With
    MsgBox "The booking details successfully updated.", vbInformation + vbOKOnly, strSystem
    
    Else
        MsgBox "Please reenter the ACQUISTION NUMBER of valid material", _
            vbExclamation
    End If
    GoTo End_Sub

Error_Handler:
    DisplayError "Database Error"
End_Sub:
End Sub


Private Sub cmd3Cancel_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    CenterForm Me, frmMDIMainMenu
    gtTableActive = "RESERVE"
    dbRetrieve lstPatronNum, "PATRON", "PATRON_NUM"
    dbRetrieveMaterialBorrowed
End Sub


Private Sub lstAcqNum_DblClick()
    txtinfo(0) = lstAcqNum
End Sub


Private Sub txtinfo_Change(Index As Integer)
    On Error GoTo Error_Handler
    Select Case Index
        Case 0                ' Show correspond TITLE from ACQUISTION NUMBER
        srchLikeLstItem Me, lstAcqNum, txtinfo(0).Text


        If Len(txtinfo(0)) Then
            SQL = "SELECT TITLE FROM MATERIAL WHERE ACQUISITION_NUM = '" _
                & txtinfo(0) & "'"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            If gLibSnap.EOF Then
                lblInfo(0) = ""
            Else
                lblInfo(0) = gLibSnap.Fields("TITLE")
            End If
        Else
            lblInfo(0) = ""
        End If
    Case 1                ' Show correspond NAME from PATRON NUMBER
        srchLikeLstItem Me, lstPatronNum, txtinfo(1).Text
        If Len(txtinfo(1)) Then
            SQL = "SELECT PATRON.NAME FROM PATRON WHERE PATRON_NUM = '" _
                & txtinfo(1) & "'"
            Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
            If gLibSnap.EOF Then
                lblInfo(1) = ""
            Else
                If Not IsNull(gLibSnap.Fields("NAME")) Then lblInfo(1) = gLibSnap.Fields("NAME")
            End If
        End If
    End Select
Exit Sub
Error_Handler:
    DisplayError "Database Error"
End Sub


Private Sub dbRetrieveMaterialBorrowed()
    On Local Error GoTo Error_Handler
    Dim SQL As Variant
    lstAcqNum.Clear
    lstAcqNum.AddItem ""
    SQL = "SELECT DISTINCT LOAN.ACQUISITION_NUM, LOAN.DATE_RETURN FROM LOAN"
    Set gLibSnap = gLibDB.OpenRecordset(SQL, dbOpenSnapshot)
    While Not gLibSnap.EOF
        If Not IsNull(gLibSnap.Fields(0)) Then
            ' Filter unRETURNED items only
            If IsNull(gLibSnap.Fields(1)) Then _
                lstAcqNum.AddItem gLibSnap.Fields(0)
        End If
        gLibSnap.MoveNext
    Wend
    If lstAcqNum.ListCount Then lstAcqNum.ListIndex = 0
    GoTo End_Sub
Error_Handler:
    DisplayError ""
Resume End_Sub
End_Sub:
End Sub


Private Sub lstPatronNum_DblClick()
    txtinfo(1) = lstPatronNum
End Sub

