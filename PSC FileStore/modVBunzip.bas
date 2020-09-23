Attribute VB_Name = "modVBunzip"
Option Explicit
'__
'__The DLL original name, unzip32.dll, was renamed to Info552-unzip32vc.dll.
'__Info = Info-ZIP  ftp://ftp.info-zip.org/pub/infozip/WIN32/unz552dN.Zip
'__552 = Version 5.52
'__unzip32 = Original name
'__vc = My mark so I know I did it (George Inacio).
'__The rename was done so will not clash with other versions of same DLL.
'__
'00-- Please Do Not Remove These Comment Lines!
'00----------------------------------------------------------------
'00-- Sample VB 5 / VB 6 code to drive Info552-unzip32vc.dll
'00-- Contributed to the Info-ZIP project by Mike Le Voi
'00--
'00-- Contact me at: mlevoi@modemss.brisnet.org.au
'00--
'00-- Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
'00--
'00-- Use this code at your own risk. Nothing implied or warranted
'00-- to work on your machine :-)
'00----------------------------------------------------------------
'00--
'00-- This Source Code Is Freely Available From The Info-ZIP Project
'00-- Web Server At:
'00-- ftp://ftp.info-zip.org/pub/infozip/infozip.html
'00--
'00-- A Very Special Thanks To Mr. Mike Le Voi
'00-- And Mr. Mike White
'00-- And The Fine People Of The Info-ZIP Group
'00-- For Letting Me Use And Modify Their Original
'00-- Visual Basic 5.0 Code! Thank You Mike Le Voi.
'00-- For Your Hard Work In Helping Me Get This To Work!!!
'00---------------------------------------------------------------
'00--
'00-- Contributed To The Info-ZIP Project By Raymond L. King.
'00-- Modified June 21, 1998
'00-- By Raymond L. King
'00-- Custom Software Designers
'00--
'00-- Contact Me At: king@ntplx.net
'00-- ICQ 434355
'00-- Or Visit Our Home Page At: http://www.ntplx.net/~king
'00--
'00---------------------------------------------------------------
'00--
'00-- Modified August 17, 1998
'00-- by Christian Spieler
'00-- (implemented sort of a "real" user interface)
'00-- Modified May 11, 2003
'00-- by Christian Spieler
'00-- (use late binding for referencing the common dialog)
'00--
'00---------------------------------------------------------------

Private Type UNZIPnames '01-- C Style argv
    uzFiles(0 To 99) As String
End Type

Private Type UNZIPCBChar '02-- Callback Large "String"
    ch(32800) As Byte
End Type

Private Type UNZIPCBCh '03-- Callback Small "String"
    ch(256) As Byte
End Type

Private Type DCLIST '04-- Info552-unzip32vc.dll DCL Structure
    ExtractOnlyNewer  As Long    ' 1 = Extract Only Newer/New, Else 0
    SpaceToUnderscore As Long    ' 1 = Convert Space To Underscore, Else 0
    PromptToOverwrite As Long    ' 1 = Prompt To Overwrite Required, Else 0
    fQuiet            As Long    ' 2 = No Messages, 1 = Less, 0 = All
    ncflag            As Long    ' 1 = Write To Stdout, Else 0
    ntflag            As Long    ' 1 = Test Zip File, Else 0
    nvflag            As Long    ' 0 = Extract, 1 = List Zip Contents
    nfflag            As Long    ' 1 = Extract Only Newer Over Existing, Else 0
    nzflag            As Long    ' 1 = Display Zip File Comment, Else 0
    ndflag            As Long    ' 1 = Honor Directories, Else 0
    noflag            As Long    ' 1 = Overwrite Files, Else 0
    naflag            As Long    ' 1 = Convert CR To CRLF, Else 0
    nZIflag           As Long    ' 1 = Zip Info Verbose, Else 0
    C_flag            As Long    ' 1 = Case Insensitivity, 0 = Case Sensitivity
    fPrivilege        As Long    ' 1 = ACL, 2 = Privileges
    Zip               As String  ' The Zip Filename To Extract Files
    ExtractDir        As String  ' The Extraction Directory, NULL If Extracting To Current Dir
End Type

Private Type USERFUNCTION '05-- Info552-unzip32vc.dll Userfunctions Structure
    UZDLLPrnt     As Long     ' Pointer To Apps Print Function
    UZDLLSND      As Long     ' Pointer To Apps Sound Function
    UZDLLREPLACE  As Long     ' Pointer To Apps Replace Function
    UZDLLPASSWORD As Long     ' Pointer To Apps Password Function
    UZDLLMESSAGE  As Long     ' Pointer To Apps Message Function
    UZDLLSERVICE  As Long     ' Pointer To Apps Service Function (Not Coded!)
    TotalSizeComp As Long     ' Total Size Of Zip Archive
    TotalSize     As Long     ' Total Size Of All Files In Archive
    CompFactor    As Long     ' Compression Factor
    NumMembers    As Long     ' Total Number Of All Files In The Archive
    cchComment    As Integer  ' Flag If Archive Has A Comment!
End Type

Private Type UZPVER '06-- Info552-unzip32vc.dll Version Structure
    structlen       As Long         ' Length Of The Structure Being Passed
    flag            As Long         ' Bit 0: is_beta  bit 1: uses_zlib
    beta            As String * 10  ' e.g., "g BETA" or ""
    Date            As String * 20  ' e.g., "4 Sep 95" (beta) or "4 September 1995"
    zlib            As String * 10  ' e.g., "1.0.5" or NULL
    unzip(1 To 4)   As Byte         ' Version Type Unzip
    zipinfo(1 To 4) As Byte         ' Version Type Zip Info
    os2dll          As Long         ' Version Type OS2 DLL
    windll(1 To 4)  As Byte         ' Version Type Windows DLL
End Type

'07-- This Assumes Info552-unzip32vc.dll Is In Your \Windows\System Directory!
'Private Declare Function Wiz_SingleEntryUnzip Lib "Info552-unzip32vc.dll" _
         (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
         ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
         dcll As DCLIST, Userf As USERFUNCTION) As Long '07
Private Declare Function Wiz_SingleEntryUnzip Lib "Info552-unzip32vc.dll" (ByVal ifnc As Long, _
                                                                           ByRef ifnv As UNZIPnames, _
                                                                           ByVal xfnc As Long, _
                                                                           ByRef xfnv As UNZIPnames, _
                                                                           dcll As DCLIST, _
                                                                           Userf As USERFUNCTION) As Long    '07

Private Declare Sub UzpVersion2 Lib "Info552-unzip32vc.dll" (uzpv As UZPVER) '07

Private UZDCL  As DCLIST '08-- Private Variables For Structure Access
Private UZUSER As USERFUNCTION '08-- Private Variables For Structure Access
Private UZVER  As UZPVER '08-- Private Variables For Structure Access

'09-- Public Variables For Setting The
'09-- Info552-unzip32vc.dll DCLIST Structure
'09-- These Must Be Set Before The Actual Call To Call_VBUnZip32
'Public uExtractOnlyNewer As Integer  ' 1 = Extract Only Newer/New, Else 0  '09
Private uExtractOnlyNewer    As Integer
'Public uSpaceUnderScore  As Integer  ' 1 = Convert Space To Underscore, Else 0  '09
Private uSpaceUnderScore As Integer
Public uPromptOverWrite  As Integer  ' 1 = Prompt To Overwrite Required, Else 0  '09
'Public uQuiet            As Integer  ' 2 = No Messages, 1 = Less, 0 = All  '09
Private uQuiet As Integer
'Public uWriteStdOut      As Integer  ' 1 = Write To Stdout, Else 0  '09
Private uWriteStdOut As Integer
'Public uTestZip          As Integer  ' 1 = Test Zip File, Else 0  '09
Private uTestZip As Integer
Public uExtractList As Integer  ' 0 = Extract, 1 = List Contents  '09
'Public uFreshenExisting  As Integer  ' 1 = Update Existing by Newer, Else 0  '09
Private uFreshenExisting As Integer
Public uDisplayComment   As Integer  ' 1 = Display Zip File Comment, Else 0  '09
Public uHonorDirectories As Integer  ' 1 = Honor Directories, Else 0  '09
Public uOverWriteFiles   As Integer  ' 1 = Overwrite Files, Else 0  '09
'Public uConvertCR_CRLF   As Integer  ' 1 = Convert CR To CRLF, Else 0  '09
Private uConvertCR_CRLF As Integer
'Public uVerbose          As Integer  ' 1 = Zip Info Verbose  '09
Private uVerbose As Integer
'Public uCaseSensitivity  As Integer  ' 1 = Case Insensitivity, 0 = Case Sensitivity  '09
Private uCaseSensitivity     As Integer
'Public uPrivilege        As Integer  ' 1 = ACL, 2 = Privileges, Else 0  '09
Private uPrivilege           As Integer
Public uZipFileName      As String   ' The Zip File Name  '09
Public uExtractDir       As String   ' Extraction Directory, Null If Current Directory  '09

'10-- Public Program Variables
Public uZipNumber    As Long         ' Zip File Number  '10
Public uNumberFiles  As Long         ' Number Of Files  '10
Public uNumberXFiles As Long         ' Number Of Extracted Files  '10
'Public uZipMessage   As String       ' For Zip Message  '10
Private uZipMessage          As String
Public uZipInfo      As String       ' For Zip Information  '10
Public uZipNames     As UNZIPnames   ' Names Of Files To Unzip  '10
Public uExcludeNames As UNZIPnames   ' Names Of Zip Files To Exclude  '10
'Public uVbSkip       As Integer      ' For DLL Password Function  '10
Private uVbSkip As Integer

Public Sub Call_VBUnZip32()

  '18-- Main Info552-unzip32vc.dll UnZip32 Subroutine
  '18-- (WARNING!) Do Not Change!

  Dim retcode As Long
  Dim MsgStr As String

  '-- Set The Info552-unzip32vc.dll Options
  '-- (WARNING!) Do Not Change

    UZDCL.ExtractOnlyNewer = uExtractOnlyNewer ' 1 = Extract Only Newer/New
    UZDCL.SpaceToUnderscore = uSpaceUnderScore ' 1 = Convert Space To Underscore
    UZDCL.PromptToOverwrite = uPromptOverWrite ' 1 = Prompt To Overwrite Required
    UZDCL.fQuiet = uQuiet                      ' 2 = No Messages 1 = Less 0 = All
    UZDCL.ncflag = uWriteStdOut                ' 1 = Write To Stdout
    UZDCL.ntflag = uTestZip                    ' 1 = Test Zip File
    UZDCL.nvflag = uExtractList                ' 0 = Extract 1 = List Contents
    UZDCL.nfflag = uFreshenExisting            ' 1 = Update Existing by Newer
    UZDCL.nzflag = uDisplayComment             ' 1 = Display Zip File Comment
    UZDCL.ndflag = uHonorDirectories           ' 1 = Honour Directories
    UZDCL.noflag = uOverWriteFiles             ' 1 = Overwrite Files
    UZDCL.naflag = uConvertCR_CRLF             ' 1 = Convert CR To CRLF
    UZDCL.nZIflag = uVerbose                   ' 1 = Zip Info Verbose
    UZDCL.C_flag = uCaseSensitivity            ' 1 = Case insensitivity, 0 = Case Sensitivity
    UZDCL.fPrivilege = uPrivilege              ' 1 = ACL 2 = Priv
    UZDCL.Zip = uZipFileName                   ' ZIP Filename
    UZDCL.ExtractDir = uExtractDir             ' Extraction Directory, NULL If Extracting
    ' To Current Directory

    '-- Set Callback Addresses
    '-- (WARNING!!!) Do Not Change
    With UZUSER
        .UZDLLPrnt = FnPtr(AddressOf UZDLLPrnt)
        .UZDLLSND = 0&    '-- Not Supported
        .UZDLLREPLACE = FnPtr(AddressOf UZDLLRep)
        .UZDLLPASSWORD = FnPtr(AddressOf UZDLLPass)
        .UZDLLMESSAGE = FnPtr(AddressOf UZReceiveDLLMessage)
        .UZDLLSERVICE = FnPtr(AddressOf UZDLLServ)
    End With 'UZUSER
    '-- Set Info552-unzip32vc.dll Version Space
    '-- (WARNING!!!) Do Not Change
    With UZVER
        .structlen = Len(UZVER)
        .beta = Space$(9) & vbNullChar
        .Date = Space$(19) & vbNullChar
        .zlib = Space$(9) & vbNullChar
    End With 'UZVER

    '-- Get Version
    UzpVersion2 UZVER
    '--------------------------------------
    '-- You Can Change This For Displaying
    '-- The Version Information!
    '--------------------------------------
    With UZVER
        MsgStr = "DLL Date: " & szTrim(.Date)
        MsgStr = MsgStr & vbNewLine & "Zip Info: " & Hex$(.zipinfo(1)) & "." & _
                 Hex$(.zipinfo(2)) & Hex$(.zipinfo(3))
        MsgStr = MsgStr & vbNewLine & "DLL Version: " & Hex$(.windll(1)) & "." & _
                 Hex$(.windll(2)) & Hex$(.windll(3))
    End With 'UZVER
    MsgStr$ = MsgStr$ & vbNewLine$ & "--------------"
    '-- End Of Version Information.

    '-- Go UnZip The Files! (Do Not Change Below!!!)
    '-- This Is The Actual UnZip Routine
    retcode = Wiz_SingleEntryUnzip(uNumberFiles, uZipNames, uNumberXFiles, _
              uExcludeNames, UZDCL, UZUSER)
    '---------------------------------------------------------------

    '-- If There Is An Error Display A MsgBox!
    If retcode <> 0 Then
        'MsgBox retcode '< Old code
        'Make the variable True so can report the error
        gbUnzipError = True
    End If
    '-- You Can Change This As Needed!
    '-- For Compression Information
    MsgStr$ = MsgStr$ & vbNewLine & "Only Shows If uExtractList = 1 List Contents"
    MsgStr$ = MsgStr$ & vbNewLine & "--------------"
    With UZUSER
        MsgStr = MsgStr & vbNewLine & "Comment         : " & .cchComment
        MsgStr = MsgStr & vbNewLine & "Total Size Comp : " & .TotalSizeComp
        MsgStr = MsgStr & vbNewLine & "Total Size      : " & .TotalSize
        MsgStr = MsgStr & vbNewLine & "Compress Factor : %" & .CompFactor
        MsgStr = MsgStr & vbNewLine & "Num Of Members  : " & .NumMembers
    End With 'UZUSER
    MsgStr$ = MsgStr$ & vbNewLine & "--------------"

    'VBUnzFrm.txtMsgOut.Text = VBUnzFrm.txtMsgOut.Text & MsgStr$ & vbNewLine

End Sub

'Public Function FnPtr(ByVal lp As Long) As Long
'SUGGESTION: Routine's name is Duplicated in another module.
'If they are identical, pick one and delete the other, otherwise rename one of them.
''11-- Puts A Function Pointer In A Structure
''11-- For Callbacks.
'    FnPtr = lp
'End Function

Public Function szTrim(szString As String) As String

  '17-- ASCIIZ To String Function

  Dim pos As Long

    pos = InStr(szString, vbNullChar)

    Select Case pos
      Case Is > 1
        szTrim = Trim$(Left$(szString, pos - 1))
      Case 1
        szTrim = vbNullString
      Case Else
        szTrim = Trim$(szString)
    End Select

End Function

Public Function UZDLLPass(ByRef p As UNZIPCBCh, _
                          ByVal n As Long, ByRef m As UNZIPCBCh, _
                          ByRef Name As UNZIPCBCh) As Integer

  '15-- Callback For Info552-unzip32vc.dll - Password Function

  Dim prompt     As String
  Dim XX         As Integer
  Dim szpassword As String

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        UZDLLPass = 1

        If uVbSkip = 1 Then
            Exit Function '---> Bottom
        End If

        '-- Get The Zip File Password
        szpassword = InputBox("Please Enter The Password!")

        '-- No Password So Exit The Function
        If Len(szpassword) = 0 Then
            uVbSkip = 1
            Exit Function '---> Bottom
        End If

        '-- Zip File Password So Process It
        For XX = 0 To 255
            If m.ch(XX) = 0 Then
                Exit For 'loop varying xx
              Else 'NOT M.CH(XX)...
                prompt = prompt & Chr$(m.ch(XX))
            End If
        Next XX

        For XX = 0 To n - 1
            p.ch(XX) = 0
        Next XX

        For XX = 0 To Len(szpassword) - 1
            p.ch(XX) = Asc(Mid$(szpassword, XX + 1, 1))
        Next XX

        p.ch(XX) = 0 ' Put Null Terminator For C

        UZDLLPass = 0
    On Error GoTo 0

End Function

Public Function UZDLLPrnt(ByRef fname As UNZIPCBChar, ByVal X As Long) As Long

  '13-- Callback For Info552-unzip32vc.dll - Print Message Function

  Dim s0 As String
  Dim XX As Long

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        s0 = vbNullString

        '-- Gets The Info552-unzip32vc.dll Message For Displaying.
        For XX = 0 To X - 1
            If fname.ch(XX) = 0 Then
                Exit For 'loop varying xx
            End If
            s0 = s0 & Chr$(fname.ch(XX))
        Next XX

        '-- Assign Zip Information
        If Mid$(s0, 1, 1) = vbLf Then
            s0 = vbNewLine ' Damn UNIX :-)
        End If
        uZipInfo = uZipInfo & s0

        UZDLLPrnt = 0
    On Error GoTo 0

End Function

Public Function UZDLLRep(ByRef fname As UNZIPCBChar) As Long

  '16-- Callback For Info552-unzip32vc.dll - Report Function To Overwrite Files.
  '16-- This Function Will Display A MsgBox Asking The User
  '16-- If They Would Like To Overwrite The Files.

  Dim s0 As String
  Dim XX As Long

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        UZDLLRep = 100 ' 100 = Do Not Overwrite - Keep Asking User
        s0 = vbNullString

        For XX = 0 To 255
            If fname.ch(XX) = 0 Then
                Exit For 'loop varying xx
            End If
            s0 = s0 & Chr$(fname.ch(XX))
        Next XX

        '-- This Is The MsgBox Code
        XX = MsgBox("Overwrite " & s0 & "?", vbExclamation & vbYesNoCancel, _
             "Call_VBUnZip32 - File Already Exists!")

        If XX = vbNo Then
            Exit Function '---> Bottom
        End If

        If XX = vbCancel Then
            UZDLLRep = 104       ' 104 = Overwrite None
            Exit Function '---> Bottom
        End If

        UZDLLRep = 102         ' 102 = Overwrite, 103 = Overwrite All
    On Error GoTo 0

End Function

Public Function UZDLLServ(ByRef mname As UNZIPCBChar, ByVal X As Long) As Long

  '14-- Callback For Info552-unzip32vc.dll - DLL Service Function

  Dim s0 As String
  Dim XX As Long

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        ' Parameter x contains the size of the extracted archive entry.
        ' This information may be used for some kind of progress display...

        s0 = vbNullString
        '-- Get Info231-zip32vc.dll Message For processing
        For XX = 0 To UBound(mname.ch)
            If mname.ch(XX) = 0 Then
                Exit For 'loop varying xx
            End If
            s0 = s0 & Chr$(mname.ch(XX))
        Next XX
        ' At this point, s0 contains the message passed from the DLL
        ' It is up to the developer to code something useful here :)

        UZDLLServ = 0 ' Setting this to 1 will abort the zip!
    On Error GoTo 0

End Function

Public Sub UZReceiveDLLMessage(ByVal ucsize As Long, _
                               ByVal csiz As Long, _
                               ByVal cfactor As Integer, _
                               ByVal mo As Integer, _
                               ByVal dy As Integer, _
                               ByVal yr As Integer, _
                               ByVal hh As Integer, _
                               ByVal mm As Integer, _
                               ByVal c As Byte, ByRef fname As UNZIPCBCh, _
                               ByRef meth As UNZIPCBCh, ByVal crc As Long, _
                               ByVal fCrypt As Byte)

  '12-- Callback For Info552-unzip32vc.dll - Receive Message Function

  Dim s0     As String
  Dim XX     As Long
  Dim strout As String * 80

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        '------------------------------------------------
        '-- This Is Where The Received Messages Are
        '-- Printed Out And Displayed.
        '-- You Can Modify Below!
        '------------------------------------------------

        strout = Space$(80)

        '-- For Zip Message Printing
        If uZipNumber = 0 Then
            Mid$(strout, 1, 50) = "Filename:"
            Mid$(strout, 53, 4) = "Size"
            Mid$(strout, 62, 4) = "Date"
            Mid$(strout, 71, 4) = "Time"
            uZipMessage = strout & vbNewLine
            strout = Space$(80)
        End If

        s0 = vbNullString

        '-- Do Not Change This For Next!!!
        For XX = 0 To 255
            If fname.ch(XX) = 0 Then
                Exit For 'loop varying xx
            End If
            s0 = s0 & Chr$(fname.ch(XX))
        Next XX

        '-- Assign Zip Information For Printing
        Mid$(strout, 1, 50) = Mid$(s0, 1, 50)
        Mid$(strout, 51, 7) = Right$("        " & CStr(ucsize), 7)
        Mid$(strout, 60, 3) = Right$("0" & Trim$(CStr(mo)), 2) & "/"
        Mid$(strout, 63, 3) = Right$("0" & Trim$(CStr(dy)), 2) & "/"
        Mid$(strout, 66, 2) = Right$("0" & Trim$(CStr(yr)), 2)
        Mid$(strout, 70, 3) = Right$(Str$(hh), 2) & ":"
        Mid$(strout, 73, 2) = Right$("0" & Trim$(CStr(mm)), 2)

        ' Mid$(strout, 75, 2) = Right$(" " & CStr(cfactor), 2)
        ' Mid$(strout, 78, 8) = Right$("        " & CStr(csiz), 8)
        ' s0 = ""
        ' For xx = 0 To 255
        '     If meth.ch(xx) = 0 Then Exit For
        '     s0 = s0 & Chr$(meth.ch(xx))
        ' Next xx

        '-- Do Not Modify Below!!!
        uZipMessage = uZipMessage & strout & vbNewLine
        uZipNumber = uZipNumber + 1
    On Error GoTo 0

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:23) 156 + 321 = 477 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 172  Code: 356  Total: 528 Lines
':) CommentOnly: 142 (26.9%)  Commented: 90 (17%)  Empty: 90 (17%)  Max Logic Depth: 3
