Attribute VB_Name = "WinAPI"



Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Win32 API declaration to get the current user of Windows
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
' Win32 API to launch file with associated application. E.g (.htm|IE, .txt|Notepad) etc
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateStatusWindow Lib "comctl32.dll" Alias "CreateStatusWindowA" (ByVal style As Long, ByVal lpszText As String, ByVal hWndParent As Long, ByVal wID As Long) As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Const LVS_EX_FULLROWSELECT = &H20

Public Const SW_SHOWNORMAL = 1


Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Const LVSCW_AUTOSIZE = -1
Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Sub AdjustColumnWidth(LV As ListView, AccountForHeaders As Boolean)
    Dim col As Integer, lParam As Long
    If AccountForHeaders Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If
    ' Send the message to all the columns
    For col = 0 To LV.ColumnHeaders.Count - 1
        SendMessage LV.hwnd, LVM_SETCOLUMNWIDTH, ByVal col, ByVal lParam
    Next
End Sub
