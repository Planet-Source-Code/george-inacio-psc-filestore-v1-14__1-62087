Attribute VB_Name = "modBrowseFolder"
Option Explicit

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Const BIF_RETURNONLYFSDIRS As Long = 1
Public Const MAX_PATH As Long = 260
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                 ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

':)Code Fixer V3.0.9 (04/08/2005 18:02:28) 23 + 0 = 23 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 24  Code: 0  Total: 24 Lines
':) CommentOnly: 1 (4.2%)  Commented: 0 (0%)  Empty: 2 (8.3%)  Max Logic Depth: 1
