Attribute VB_Name = "modVC"
Option Explicit

Public gsLocalForm As String
Public gsProgName As String
Public gsOwner As String
Public gsProgVer As String
Public DB1 As Database
Public glFormHeight As Long
Public glFormLeft As Long
Public glFormTop As Long
Public glFormWidth As Long
Public gbFormT As Boolean
Public gbUnzipError As Boolean

'Start: Scroll Horizontal List Box
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Long, _
       lParam As Any) As Long 'Scroll Horizontal List Box

Private Declare Function GetFocus Lib "user32" () As Long 'Scroll Horizontal List Box

Private Const LB_SETHORIZONTALEXTENT As Long = &H194 'Scroll Horizontal List Box
'Public Const NUL = 0& 'Scroll Horizontal List Box
Private Const NUL As Long = 0 'Scroll Horizontal List Box
'End: Scroll Horizontal List Box
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Call_DoIsRunning()

    MsgBox gsProgName & "  " & gsProgVer & _
           vbNewLine & vbNewLine & "Only One instance of this program may run at one time." & _
           vbNewLine & "If you don't see the program on taskbar then most probably" & _
           vbNewLine & "you need to restart the system."

End Sub

Public Sub Call_MBAlreadyExists(ByVal sSearch As String, ByVal sTheName As String)

    MsgBox sSearch _
           & vbNewLine & sTheName & " already exists!" _
           & vbNewLine & "Please Try A Different " & sTheName & " !" _
           & vbNewLine & "Click OK Button To Return.", _
           vbOKOnly + vbCritical, gsLocalForm

End Sub

Public Function Func_BoxBlank(ByVal sBoxBlank As String) As Boolean

    MsgBox sBoxBlank & " can NOT be Blank!" & _
           vbNewLine & "Please Try Again!", vbOKOnly + vbCritical, gsLocalForm
    Func_BoxBlank = True

End Function

Public Function Func_ClearChar0To31(ByVal sStringToFix As String) As String

  Dim lCountString As Long
  Dim sNewString As String
  Dim sOneChar As String

  'Removes characters 0 to 31 and 127 from the string
  'These characters have no graphical representation
  'And can create problems if used on directories or file names.

    lCountString = Len(sStringToFix)
    For lCountString = 1 To lCountString
        If Asc(Mid$(sStringToFix, lCountString, 1)) <= 31 Then
            sOneChar = vbNullString
          Else 'NOT ASC(MID$(SSTRINGTOFIX,...
            If Asc(Mid$(sStringToFix, lCountString, 1)) = 127 Then
                sOneChar = vbNullString
              Else 'NOT ASC(MID$(SSTRINGTOFIX,...
                sOneChar = Mid$(sStringToFix, lCountString, 1)
            End If
        End If
        sNewString = sNewString & sOneChar
    Next lCountString
    Func_ClearChar0To31 = sNewString

End Function

Public Function Func_FilterString(ByVal sStringToFix As String) As String

    sStringToFix = Replace$(sStringToFix, "*", vbNullString)
    sStringToFix = Replace$(sStringToFix, "<", vbNullString)
    sStringToFix = Replace$(sStringToFix, ">", vbNullString)
    sStringToFix = Replace$(sStringToFix, "?", vbNullString)
    sStringToFix = Replace$(sStringToFix, ":", vbNullString)
    sStringToFix = Replace$(sStringToFix, "|", vbNullString)
    sStringToFix = Replace$(sStringToFix, "  ", " ")
    sStringToFix = Replace$(sStringToFix, "/", "-")
    sStringToFix = Replace$(sStringToFix, "\", "-")
    sStringToFix = Replace$(sStringToFix, ".", "_")

    Func_FilterString = sStringToFix

End Function

Public Function Func_MaxBoxLength(ByVal sBoxText As String, ByVal sMessage As String, _
                                  ByRef lMaxLength As Long) As Boolean

  Dim lHowMany As Long
  Dim sPlural As String

    If Len(Trim$(sBoxText)) > lMaxLength Then
        Func_MaxBoxLength = True
        lHowMany = Len(sBoxText) - lMaxLength
        If lHowMany > 1 Then
            sPlural = " characters."
          Else 'NOT LHOWMANY...
            sPlural = " character."
        End If
        MsgBox "(" & sBoxText & ")" _
                 & vbNewLine & "Maximum characters for " & sMessage _
                 & " is " & UCase$(Func_NumToText(lMaxLength)) _
                 & " (" & lMaxLength & ")!" _
                 & vbNewLine & "The Box contains " & Len(sBoxText) & " characters. Remove " _
                 & lHowMany & sPlural _
                 & vbNewLine & "Please Try A Different One!" _
                 & vbNewLine & "Click OK Button To Return.", _
                 vbOKOnly + vbCritical, gsLocalForm
    End If

End Function

Public Function Func_NumToText(ByVal sPassNumber As String) As String

  Dim sZeroNineteen(19) As String
  Dim sTwentyNinety(8) As String
  Dim Formated As String
  Dim Hun As Long
  Dim Tens As Long
  'I copy this function from some where but I do not remember the author name
  'If you recognise it please let me know so I can place your name here
  'And thanks for the function
  'THIS FUNCTION CONVERTS 1 TO 999 INTO WORDS

    sZeroNineteen(0) = vbNullString '""
    sZeroNineteen(1) = "one"
    sZeroNineteen(2) = "two"
    sZeroNineteen(3) = "three"
    sZeroNineteen(4) = "four"
    sZeroNineteen(5) = "five"
    sZeroNineteen(6) = "six"
    sZeroNineteen(7) = "seven"
    sZeroNineteen(8) = "eight"
    sZeroNineteen(9) = "nine"
    sZeroNineteen(10) = "ten"
    sZeroNineteen(11) = "eleven"
    sZeroNineteen(12) = "twelve"
    sZeroNineteen(13) = "thirteen"
    sZeroNineteen(14) = "fourteen"
    sZeroNineteen(15) = "fifteen"
    sZeroNineteen(16) = "sixteen"
    sZeroNineteen(17) = "seventeen"
    sZeroNineteen(18) = "eighteen"
    sZeroNineteen(19) = "nineteen"
    sTwentyNinety(0) = "twenty"
    sTwentyNinety(1) = "thirty"
    sTwentyNinety(2) = "forty"
    sTwentyNinety(3) = "fifty"
    sTwentyNinety(4) = "sixty"
    sTwentyNinety(5) = "seventy"
    sTwentyNinety(6) = "eighty"
    sTwentyNinety(7) = "ninety"
    Formated = Format$(sPassNumber, "000.00")
    Hun = Mid$(Formated, 1, 1)
    Tens = Mid$(Formated, 2, 2)
    If Hun <> 0 Then
        Func_NumToText = sZeroNineteen(Hun) & " hundred and "
    End If
    If Tens <> 0 Then
        If Tens < 20 Then
            Func_NumToText = Func_NumToText + sZeroNineteen(Tens) & " "
          Else 'NOT TENS...
            Func_NumToText = Func_NumToText + sTwentyNinety(Mid$(Tens, 1, 1) - 2) & " " & sZeroNineteen(Mid$(Tens, 2, 1)) & " "
        End If
    End If

End Function

Public Function Func_SrchReplace(ByVal sStringToFix As String) As String

    sStringToFix = Replace$(sStringToFix, Chr$(39), "`")
    sStringToFix = Replace$(sStringToFix, Chr$(34), "``")
    Func_SrchReplace = sStringToFix

End Function

':)Code Fixer V3.0.9 (04/08/2005 18:02:24) 20 + 146 = 166 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 30  Code: 167  Total: 197 Lines
':) CommentOnly: 11 (5.6%)  Commented: 9 (4.6%)  Empty: 34 (17.3%)  Max Logic Depth: 4
