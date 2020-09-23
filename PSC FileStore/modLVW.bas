Attribute VB_Name = "modLVW"
Option Explicit
'ListView Autoresize

'Private Const LVM_FIRST As Long = &H1000
'Public Const LVM_SETCOLUMNWIDTH As Double = (LVM_FIRST + 30)
'Public Const LVSCW_AUTOSIZE As Long = -1
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                                                       ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 12  Code: 0  Total: 12 Lines
':) CommentOnly: 5 (41.7%)  Commented: 0 (0%)  Empty: 1 (8.3%)  Max Logic Depth: 0
