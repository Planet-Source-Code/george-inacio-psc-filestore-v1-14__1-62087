VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Code done by Ulli
'Custom Tooltip Class
'''''''''''''''''''''
'This class was inspired by code by Eidos (found at PSC some time ago) and others.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jan02 2003   UMG
'
'Three new options have been added - display tooltip always / only if parent form is active / None
'see TTStyle.
'
'Added missing Style private property variable.
'Rearranged code a little.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private tthWnd                             As Long           'Tooltip window handle
Private Const ToolTipWindowClassName       As String = "Tooltips_Class32"
Private Const CW_USEDEFAULT                As Long = &H80000000
Private Const TTS_STANDARD                 As Long = 0
Private Const TTS_BALLOON                  As Long = &H40
Private Const TTS_ALWAYSTIP                As Long = 1       'display even if parent window is inactive
Private Const TTS_NOPREFIX                 As Long = 2       'does not remove "&" from text
Private Const TTDT_AUTOPOP                 As Long = 2
Private Const TTDT_INITIAL                 As Long = 3
Public Enum TTStyle
    TTStandardIfActive = TTS_STANDARD                   'suppress if parent form is not active
    TTBalloonIfActive = TTS_BALLOON                     'suppress if parent form is not active
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP    'display even if parent form is not active
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP      'display even if parent form is not active
    TTNone = -1                                         'kill tooltip (this is simply treated as illegal)
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone
#End If
''''''''<:-) :SUGGESTION: Inserted by Code Fixer. (Must be placed after Enum Declaration for Code Fixer to recognize it properly)
#If False Then
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone 'to preserve the case
#End If
Public Enum TTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
    'all have a light gray shadow so be careful when selecting the ToolTip BackColor
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError
#End If
''''''''''''<:-) :SUGGESTION: Inserted by Code Fixer. (Must be placed after Enum Declaration for Code Fixer to recognize it properly)
#If False Then
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError 'to preserve the case
#End If
'my properties
Private myStyle                            As TTStyle
Private myIcon                             As TTIcon
Private myForeColor                        As Long
Private myBackColor                        As Long
Private myTitle                            As String         'has the current title
Private myHoverTime                        As Long           'time im millisecs (-1 = use default)
Private myPopupTime                        As Long           'time im millisecs (-1 = use default)
Private myInitialText                      As Variant        'has the initial text
Private myInitialTitle                     As Variant        'has the initial title
Private Const SWP_NOSIZE                   As Long = &H1
Private Const SWP_NOMOVE                   As Long = &H2
Private Const SWP_NOACTIVATE               As Long = &H10
Private Const SWP_FLAGS                    As Long = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Private Const TOPMOST                      As Long = -1
Private Const WM_USER                      As Long = &H400
Private Const TTM_SETDELAYTIME             As Long = WM_USER + 3
Private Const TTM_ADDTOOL                  As Long = WM_USER + 4
Private Const TTM_SETTIPBKCOLOR            As Long = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR          As Long = WM_USER + 20
Private Const TTM_SETTITLE                 As Long = WM_USER + 32
Private Type RECTANGLE
    Left                                     As Long
    Top                                      As Long
    Right                                    As Long
    Bottom                                   As Long
End Type
Private Type ToolInfo
    ttSize                                   As Long
    myFlags                                  As Long
    ttParhWnd                                As Long
    ttId                                     As Long
    ParentRect                               As RECTANGLE
    hInstance                                As Long
    myText                                   As String
    lParam                                   As Long
End Type
Private ToolInfo                           As ToolInfo
'tool property flag bits                             meaning
Private Const TTF_CENTERTIP                As Long = 2       'center tool on parent
Private Const TTF_SUBCLASS                 As Long = &H10    'use implicit subclassinf
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal X As Long, _
                                                                              ByVal Y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hWndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECTANGLE) As Long

Public Property Get BackCol() As Long

  'this returns the current tooltip backcolor

    BackCol = myBackColor

End Property

Public Property Get Centered() As Boolean

  'this returns the current tooltip alignment

    Centered = CBool(ToolInfo.myFlags And TTF_CENTERTIP)

End Property

Private Sub Class_Initialize()

    InitCommonControls 'doesn't matter that this is called for every class instance
    myStyle = TTNone

End Sub

Private Sub Class_Terminate()

  'kill tooltip window if one exists

    If tthWnd Then
        DestroyWindow tthWnd
        tthWnd = 0
    End If
    myStyle = TTNone

End Sub

Public Function Create(cntParent As Control, _
                       ByVal strText As String, _
                       Optional ByVal ttsStyle As TTStyle = TTBalloonAlways, _
                       Optional ByVal Centered As Boolean = False, _
                       Optional ByVal ttiIcon As TTIcon = TTIconNone, _
                       Optional ByVal Title As String = vbNullString, _
                       Optional ByVal lngForeColor As Long = vbButtonText, _
                       Optional ByVal lngBackColor As Long = vbInfoBackground, _
                       Optional ByVal HoverTime As Long = -1, _
                       Optional ByVal PopupTime As Long = 99000) As Long

    PopupTime = 8000
    lngForeColor = 0
    lngBackColor = -2147483624
    'Create the tooltip window for parent control
    'This cannot create custom tooltips for hWnd-less controls
    Class_Terminate 'kill tooltip window if one exists
    With ToolInfo
        On Error Resume Next
            .ttParhWnd = cntParent.hwnd
            If Err.Number = 0 And (ttsStyle = TTBalloonAlways Or ttsStyle = TTStandardAlways Or ttsStyle = TTBalloonIfActive Or ttsStyle = TTStandardIfActive) And (ttiIcon = TTIconError Or ttiIcon = TTIconInfo Or ttiIcon = TTIconNone Or ttiIcon = TTIconWarning) Then
                'the tooltip parent control has an hWnd and the params are acceptable
                .ttSize = Len(ToolInfo)
                .myFlags = TTF_SUBCLASS Or IIf(Centered, TTF_CENTERTIP, 0&)
                GetClientRect .ttParhWnd, .ParentRect
                .hInstance = App.hInstance
                myTitle = Title
                If myInitialTitle = Empty Then
                    myInitialTitle = myTitle
                End If
                .myText = Replace$(strText, "|", vbNewLine)       'the vertical bar is used as line break character
                If Len(myTitle) = 0 Then
                    .myText = Replace$(.myText, vbNewLine, " ")
                End If
                If myInitialText = Empty Then
                    myInitialText = .myText
                End If
                If lngForeColor < 0 Then
                    lngForeColor = GetSysColor(lngForeColor And &H7FFFFFFF) 'GetSysColor(ForeColor And &H7FFFFFFF)
                End If
                If lngBackColor < 0 Then
                    lngBackColor = GetSysColor(lngBackColor And &H7FFFFFFF)
                End If
                If lngForeColor = lngBackColor Then
                    lngForeColor = vbButtonText
                    lngBackColor = vbInfoBackground
                End If
                myForeColor = lngForeColor
                myBackColor = lngBackColor
                myStyle = ttsStyle
                myIcon = ttiIcon
                myHoverTime = HoverTime
                myPopupTime = PopupTime
                myTitle = "Tool Tip "
                'create tooltip window and set it's properties
                tthWnd = CreateWindowEx(0&, ToolTipWindowClassName, vbNullString, TTS_NOPREFIX Or ttsStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, .ttParhWnd, 0&, .hInstance, 0&)
                SetWindowPos tthWnd, TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS
                SendMessage tthWnd, TTM_ADDTOOL, 0&, ToolInfo
                SendMessage tthWnd, TTM_SETTITLE, ttiIcon, ByVal myTitle
                SendMessage tthWnd, TTM_SETTIPTEXTCOLOR, myForeColor, ByVal 0&
                SendMessage tthWnd, TTM_SETTIPBKCOLOR, myBackColor, ByVal 0&
                SendMessage tthWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal myHoverTime
                SendMessage tthWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal myPopupTime
                Create = tthWnd
            End If
        On Error GoTo 0
    End With 'ToolInfo

End Function

Public Property Get ForeCol() As Long

  'this returns the current tooltip forecolor

    ForeCol = myForeColor

End Property

Public Property Get HoverTime() As Long

  'this returns the current mouse HoverTime time in millicecs (-1 for default)

    HoverTime = myHoverTime

End Property

Public Property Get Icon() As TTIcon

  'this returns the current tooltip icon

    Icon = myIcon

End Property

Public Property Get InitialText() As String

  'this returns the inital tooltip text, ie the one that was supplied on creation

    InitialText = myInitialText

End Property

Public Property Get InitialTitle() As String

  'this returns the inital tooltip title, ie the one that was supplied on creation

    InitialTitle = myInitialTitle

End Property

Public Property Get PopupTime() As Long

  'this returns the current max PopupTime time in millisecs (-1 for default)

    PopupTime = myPopupTime

End Property

Public Property Get Style() As TTStyle

  'this returns the current tooltip style

    Style = myStyle

End Property

Public Property Get Text() As String

  'this returns the current tooltip text

    Text = ToolInfo.myText

End Property

Public Property Get Title() As String

  'this returns the current tooltip Title

    Title = myTitle

End Property

':)Code Fixer V3.0.9 (04/08/2005 18:02:30) 101 + 174 = 275 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 123  Code: 182  Total: 305 Lines
':) CommentOnly: 37 (12.1%)  Commented: 27 (8.9%)  Empty: 56 (18.4%)  Max Logic Depth: 4
