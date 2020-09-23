Attribute VB_Name = "modVBzip"
Option Explicit
'__
'__The DLL original name, zip32.dll, was renamed to Info231-zip32vc.dll.
'__Info = Info-ZIP  ftp://ftp.info-zip.org/pub/infozip/WIN32/zip231dN.Zip
'__231 = Version 2.31
'__zip32 = Original name
'__vc = My mark so I know I did it (George Inacio).
'__The rename was done so will not clash with other versions of same DLL.
'__
'00---------------------------------------------------------------
'00-- Please Do Not Remove These Comments!!!
'00---------------------------------------------------------------
'00-- Sample VB 5 code to drive Info231-zip32vc.dll
'00-- Contributed to the Info-ZIP project by Mike Le Voi
'00--
'00-- Contact me at: mlevoi@modemss.brisnet.org.au
'00--
'00-- Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
'00--
'00-- Use this code at your own risk. Nothing implied or warranted
'00-- to work on your machine :-)
'00---------------------------------------------------------------
'00--
'00-- The Source Code Is Freely Available From Info-ZIP At:
'00-- http://www.cdrom.com/pub/infozip/infozip.html
'00--
'00-- A Very Special Thanks To Mr. Mike Le Voi
'00-- And Mr. Mike White Of The Info-ZIP
'00-- For Letting Me Use And Modify His Orginal
'00-- Visual Basic 5.0 Code! Thank You Mike Le Voi.
'00---------------------------------------------------------------
'00--
'00-- Contributed To The Info-ZIP Project By Raymond L. King
'00-- Modified June 21, 1998
'00-- By Raymond L. King
'00-- Custom Software Designers
'00--
'00-- Contact Me At: king@ntplx.net
'00-- ICQ 434355
'00-- Or Visit Our Home Page At: http://www.ntplx.net/~king
'00--
'00---------------------------------------------------------------
'00
'00 This is the original example with some small changes. Only
'00 use with the original Info231-zip32vc.dll (Zip 2.3).  Do not use this VB
'00 example with Zip32z64.dll (Zip 3.0).
'00
'00 4/29/2004 Ed Gordon
'00
'00---------------------------------------------------------------
'00 Usage notes:
'00
'00 This code uses Info231-zip32vc.dll.  You DO NOT need to register the
'00 DLL to use it.  You also DO NOT need to reference it in your
'00 VB project.  You DO have to copy the DLL to your SYSTEM
'00 directory, your VB project directory, or place it in a directory
'00 on your command PATH.
'00
'00 A bug has been found in the Info231-zip32vc.dll when called from VB.  If
'00 you try to pass any values other than NULL in the ZPOPT strings
'00 Date, szRootDir, or szTempDir they get converted from the
'00 VB internal wide character format to temporary byte strings by
'00 the calling interface as they are supposed to.  However when
'00 ZpSetOptions returns the passed strings are deallocated unless the
'00 VB debugger prevents it by a break between ZpSetOptions and
'00 ZpArchive.  When Info231-zip32vc.dll uses these pointers later it
'00 can result in unpredictable behavior.  A kluge is available
'00 for Info231-zip32vc.dll, just replacing api.c in Zip 2.3, but better to just
'00 use the new Zip32z64.dll where these bugs are fixed.  However,
'00 the kluge has been added to Zip 2.31.  To determine the version
'00 of the dll you have right click on it, select the Version tab,
'00 and verify the Product Version is at least 2.31.
'00
'00 Another bug is where -R is used with some other options and can
'00 crash the dll.  This is a bug in how zip processes the command
'00 line and should be mostly fixed in Zip 2.31.  If you run into
'00 problems try using -r instead for recursion.  The bug is fixed
'00 in Zip 3.0 but note that Zip 3.0 creates dll zip32z64.dll and
'00 it is not compatible with older VB including this example.  See
'00 the new VB example code included with Zip 3.0 for calling
'00 interface changes.
'00
'00 Note that Zip32 is probably not thread safe.  It may be made
'00 thread safe in a later version, but for now only one thread in
'00 one program should use the DLL at a time.  Unlike Zip, UnZip is
'00 probably thread safe, but an exception to this has been
'00 found.  See the UnZip documentation for the latest on this.
'00
'00 All code in this VB project is provided under the Info-Zip license.
'00
'00 If you have any questions please contact Info-Zip at
'00 http://www.info-zip.org.
'00
'00 4/29/2004 EG (Updated 3/1/2005 EG)
'00
'00---------------------------------------------------------------
'00
'01-- C Style argv
'01-- Holds The Zip Archive Filenames
'01 Max for this just over 8000 as each pointer takes up 4 bytes and
'01 VB only allows 32 kB of local variables and that includes function
'01 parameters.  - 3/19/2004 EG
'01
Public Type ZIPnames '01
    zFiles(0 To 99) As String
End Type

Public Type ZipCBChar '02-- Call Back "String"
    ch(4096) As Byte
End Type

Public Type ZPOPT '03-- ZPOPT Is Used To Set The Options In The Info231-zip32vc.dll
    Date           As String ' US Date (8 Bytes Long) "12/31/98"?
    szRootDir      As String ' Root Directory Pathname (Up To 256 Bytes Long)
    szTempDir      As String ' Temp Directory Pathname (Up To 256 Bytes Long)
    fTemp          As Long   ' 1 If Temp dir Wanted, Else 0
    fSuffix        As Long   ' Include Suffixes (Not Yet Implemented!)
    fEncrypt       As Long   ' 1 If Encryption Wanted, Else 0
    fSystem        As Long   ' 1 To Include System/Hidden Files, Else 0
    fVolume        As Long   ' 1 If Storing Volume Label, Else 0
    fExtra         As Long   ' 1 If Excluding Extra Attributes, Else 0
    fNoDirEntries  As Long   ' 1 If Ignoring Directory Entries, Else 0
    fExcludeDate   As Long   ' 1 If Excluding Files Earlier Than Specified Date, Else 0
    fIncludeDate   As Long   ' 1 If Including Files Earlier Than Specified Date, Else 0
    fVerbose       As Long   ' 1 If Full Messages Wanted, Else 0
    fQuiet         As Long   ' 1 If Minimum Messages Wanted, Else 0
    fCRLF_LF       As Long   ' 1 If Translate CR/LF To LF, Else 0
    fLF_CRLF       As Long   ' 1 If Translate LF To CR/LF, Else 0
    fJunkDir       As Long   ' 1 If Junking Directory Names, Else 0
    fGrow          As Long   ' 1 If Allow Appending To Zip File, Else 0
    fForce         As Long   ' 1 If Making Entries Using DOS File Names, Else 0
    fMove          As Long   ' 1 If Deleting Files Added Or Updated, Else 0
    fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
    fUpdate        As Long   ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
    fFreshen       As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
    fJunkSFX       As Long   ' 1 If Junking SFX Prefix, Else 0
    fLatestTime    As Long   ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
    fComment       As Long   ' 1 If Putting Comment In Zip File, Else 0
    fOffsets       As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
    fPrivilege     As Long   ' 1 If Not Saving Privileges, Else 0
    fEncryption    As Long   ' Read Only Property!!!
    fRecurse       As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
    fRepair        As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
    flevel         As Byte   ' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type

Public Type ZIPUSERFUNCTIONS '04-- This Structure Is Used For The Info231-zip32vc.dll Function Callbacks
    ZDLLPrnt     As Long        ' Callback Info231-zip32vc.dll Print Function
    ZDLLCOMMENT  As Long        ' Callback Info231-zip32vc.dll Comment Function
    ZDLLPASSWORD As Long        ' Callback Info231-zip32vc.dll Password Function
    ZDLLSERVICE  As Long        ' Callback Info231-zip32vc.dll Service Function
End Type

'Public ZOPT  As ZPOPT '05-- Local Declarations
Private ZOPT             As ZPOPT
'Public ZUSER As ZIPUSERFUNCTIONS '05-- Local Declarations
Private ZUSER            As ZIPUSERFUNCTIONS

'06-- This Assumes Info231-zip32vc.dll Is In Your \Windows\System Directory!
'06-- (alternatively, a copy of Info231-zip32vc.dll needs to be located in the program
'06-- directory or in some other directory listed in PATH.)
Private Declare Function ZpInit Lib "Info231-zip32vc.dll" _
        (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long '-- Set Zip Callbacks'06

Private Declare Function ZpSetOptions Lib "Info231-zip32vc.dll" _
        (ByRef Opts As ZPOPT) As Long '-- Set Zip Options'06

Private Declare Function ZpGetOptions Lib "Info231-zip32vc.dll" _
        () As ZPOPT '-- Used To Check Encryption Flag Only'06

Private Declare Function ZpArchive Lib "Info231-zip32vc.dll" _
        (ByVal argc As Long, ByVal funame As String, _
        ByRef argv As ZIPnames) As Long '-- Real Zipping Action'06

'07-------------------------------------------------------
'07-- Public Variables For Setting The ZPOPT Structure...
'07-- (WARNING!!!) You Must Set The Options That You
'07-- Want The Info231-zip32vc.dll To Do!
'07-- Before Calling Func_VBZip32!
'07--
'07-- NOTE: See The Above ZPOPT Structure Or The Func_VBZip32
'07--       Function, For The Meaning Of These Variables
'07--       And How To Use And Set Them!!!
'07-- These Parameters Must Be Set Before The Actual Call
'07-- To The Func_VBZip32 Function!
'07-------------------------------------------------------
'Public zDate         As String '07--
Private zDate            As String
Public zRootDir      As String '07--
Public zTempDir      As String '07--
'Public zSuffix       As Integer '07--
Private zSuffix          As Integer
'Public zEncrypt      As Integer '07--
Private zEncrypt         As Integer
'Public zSystem       As Integer '07--
Private zSystem          As Integer
'Public zVolume       As Integer '07--
Private zVolume          As Integer
'Public zExtra        As Integer '07--
Private zExtra           As Integer
'Public zNoDirEntries As Integer '07--
Private zNoDirEntries    As Integer
'Public zExcludeDate  As Integer '07--
Private zExcludeDate     As Integer
'Public zIncludeDate  As Integer '07--
Private zIncludeDate     As Integer
'Public zVerbose      As Integer '07--
Private zVerbose         As Integer
'Public zQuiet        As Integer '07--
Private zQuiet           As Integer
'Public zCRLF_LF      As Integer '07--
Private zCRLF_LF         As Integer
'Public zLF_CRLF      As Integer '07--
Private zLF_CRLF         As Integer
Public zJunkDir      As Integer '07--
'Public zRecurse      As Integer '07--
Private zRecurse         As Integer
Public zGrow         As Integer '07--
'Public zForce        As Integer '07--
Private zForce           As Integer
Public zMove         As Integer '07--
'Public zDelEntries   As Integer '07--
Private zDelEntries      As Integer
'Public zUpdate       As Integer '07--
Private zUpdate          As Integer
'Public zFreshen      As Integer '07--
Private zFreshen         As Integer
'Public zJunkSFX      As Integer '07--
Private zJunkSFX         As Integer
'Public zLatestTime   As Integer '07--
Private zLatestTime      As Integer
'Public zComment      As Integer '07--
Private zComment         As Integer
'Public zOffsets      As Integer '07--
Private zOffsets         As Integer
'Public zPrivilege    As Integer '07--
Private zPrivilege       As Integer
'Public zEncryption   As Integer '07--
Private zEncryption      As Integer
'Public zRepair       As Integer '07--
Private zRepair          As Integer
Public zLevel        As Integer '07--

'08-- Public Program Variables
Public zArgc         As Integer     ' Number Of Files To Zip Up'08
Public zZipFileName  As String      ' The Zip File Name ie: Myzip.zip'08
Public zZipFileNames As ZIPnames    ' File Names To Zip Up'08
'Public zZipInfo      As String      ' Holds The Zip File Information'08

'09-- Public Constants
'09-- For Zip & UnZip Error Codes!
Private Const ZE_OK As Integer = 0              ' Success (No Error)'09
Private Const ZE_EOF As Integer = 2             ' Unexpected End Of Zip File Error'09
Private Const ZE_FORM As Integer = 3            ' Zip File Structure Error'09
Private Const ZE_MEM As Integer = 4             ' Out Of Memory Error'09
Private Const ZE_LOGIC As Integer = 5           ' Internal Logic Error'09
Private Const ZE_BIG As Integer = 6             ' Entry Too Large To Split Error'09
Private Const ZE_NOTE As Integer = 7            ' Invalid Comment Format Error'09
Private Const ZE_TEST As Integer = 8            ' Zip Test (-T) Failed Or Out Of Memory Error'09
Private Const ZE_ABORT As Integer = 9           ' User Interrupted Or Termination Error'09
Private Const ZE_TEMP As Integer = 10           ' Error Using A Temp File'09
Private Const ZE_READ As Integer = 11           ' Read Or Seek Error'09
Private Const ZE_NONE As Integer = 12           ' Nothing To Do Error'09
Private Const ZE_NAME As Integer = 13           ' Missing Or Empty Zip File Error'09
Private Const ZE_WRITE As Integer = 14          ' Error Writing To A File'09
Private Const ZE_CREAT As Integer = 15          ' Could't Open To Write Error'09
Private Const ZE_PARMS As Integer = 16          ' Bad Command Line Argument Error'09
Private Const ZE_OPEN As Integer = 18           ' Could Not Open A Specified File To Read Error'09

Public Function FnPtr(ByVal lp As Long) As Long

  '10-- These Functions Are For The Info231-zip32vc.dll
  '10--
  '10-- Puts A Function Pointer In A Structure
  '10-- For Use With Callbacks...

    FnPtr = lp

End Function

Public Function Func_VBZip32() As Long

  '15-- Main Info231-zip32vc.dll Subroutine.
  '15-- This Is Where It All Happens!!!
  '15--
  '15-- (WARNING!) Do Not Change This Function!!!
  '15--

  Dim retcode As Long

    On Error Resume Next '-- Nothing Will Go Wrong :-)

        'retcode = 0

        '-- Set Address Of Info231-zip32vc.dll Callback Functions
        '-- (WARNING!) Do Not Change!!!
        With ZUSER
            .ZDLLPrnt = FnPtr(AddressOf ZDLLPrnt)
            .ZDLLPASSWORD = FnPtr(AddressOf ZDLLPass)
            .ZDLLCOMMENT = FnPtr(AddressOf ZDLLComm)
            .ZDLLSERVICE = FnPtr(AddressOf ZDLLServ)
            '-- Set Info231-zip32vc.dll Callbacks
        End With 'ZUSER

        '-- Set Info231-zip32vc.dll Callbacks
        retcode = ZpInit(ZUSER)
        If retcode = 0 Then
            MsgBox "Info231-zip32vc.dll did not initialize.  Is it in the current directory " & _
                   "or on the command path?", vbOKOnly, "VB Zip"
            Exit Function '---> Bottom
        End If

        '-- Setup ZIP32 Options
        '-- (WARNING!) Do Not Change!
        ZOPT.Date = zDate                  ' "12/31/79"? US Date?
        ZOPT.szRootDir = zRootDir          ' Root Directory Pathname
        ZOPT.szTempDir = zTempDir          ' Temp Directory Pathname
        ZOPT.fSuffix = zSuffix             ' Include Suffixes (Not Yet Implemented)
        ZOPT.fEncrypt = zEncrypt           ' 1 If Encryption Wanted
        ZOPT.fSystem = zSystem             ' 1 To Include System/Hidden Files
        ZOPT.fVolume = zVolume             ' 1 If Storing Volume Label
        ZOPT.fExtra = zExtra               ' 1 If Including Extra Attributes
        ZOPT.fNoDirEntries = zNoDirEntries ' 1 If Ignoring Directory Entries
        ZOPT.fExcludeDate = zExcludeDate   ' 1 If Excluding Files Earlier Than A Specified Date
        ZOPT.fIncludeDate = zIncludeDate   ' 1 If Including Files Earlier Than A Specified Date
        ZOPT.fVerbose = zVerbose           ' 1 If Full Messages Wanted
        ZOPT.fQuiet = zQuiet               ' 1 If Minimum Messages Wanted
        ZOPT.fCRLF_LF = zCRLF_LF           ' 1 If Translate CR/LF To LF
        ZOPT.fLF_CRLF = zLF_CRLF           ' 1 If Translate LF To CR/LF
        ZOPT.fJunkDir = zJunkDir           ' 1 If Junking Directory Names
        ZOPT.fGrow = zGrow                 ' 1 If Allow Appending To Zip File
        ZOPT.fForce = zForce               ' 1 If Making Entries Using DOS Names
        ZOPT.fMove = zMove                 ' 1 If Deleting Files Added Or Updated
        ZOPT.fDeleteEntries = zDelEntries  ' 1 If Files Passed Have To Be Deleted
        ZOPT.fUpdate = zUpdate             ' 1 If Updating Zip File-Overwrite Only If Newer
        ZOPT.fFreshen = zFreshen           ' 1 If Freshening Zip File-Overwrite Only
        ZOPT.fJunkSFX = zJunkSFX           ' 1 If Junking SFX Prefix
        ZOPT.fLatestTime = zLatestTime     ' 1 If Setting Zip File Time To Time Of Latest File In Archive
        ZOPT.fComment = zComment           ' 1 If Putting Comment In Zip File
        ZOPT.fOffsets = zOffsets           ' 1 If Updating Archive Offsets For SFX Files
        ZOPT.fPrivilege = zPrivilege       ' 1 If Not Saving Privelages
        ZOPT.fEncryption = zEncryption     ' Read Only Property!
        ZOPT.fRecurse = zRecurse           ' 1 or 2 If Recursing Into Subdirectories
        ZOPT.fRepair = zRepair             ' 1 = Fix Archive, 2 = Try Harder To Fix
        ZOPT.flevel = zLevel               ' Compression Level - (0 To 9) Should Be 0!!!

        '-- Set Info231-zip32vc.dll Options
        retcode = ZpSetOptions(ZOPT)

        '-- Go Zip It Them Up!
        retcode = ZpArchive(zArgc, zZipFileName, zZipFileNames)

        '-- Return The Function Code
        Func_VBZip32 = retcode
    On Error GoTo 0

End Function

Public Function ZDLLComm(ByRef s1 As ZipCBChar) As Integer

  '14-- Callback For Info231-zip32vc.dll - DLL Comment Function

  Dim XX As Integer
  Dim szcomment As String

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        ZDLLComm = 1
        szcomment = InputBox("Enter the comment")
        If LenB(szcomment) = 0 Then
            Exit Function '---> Bottom
        End If
        For XX = 0 To Len(szcomment) - 1
            s1.ch(XX) = Asc(Mid$(szcomment, XX + 1, 1))
        Next XX
        s1.ch(XX) = vbNullChar ' Put null terminator for C
    On Error GoTo 0

End Function

Public Function ZDLLPass(ByRef p As ZipCBChar, _
                         ByVal n As Long, ByRef m As ZipCBChar, _
                         ByRef Name As ZipCBChar) As Integer

  '13-- Callback For Info231-zip32vc.dll - DLL Password Function

  Dim prompt     As String
  Dim XX         As Integer
  Dim szpassword As String

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        ZDLLPass = 1

        '-- If There Is A Password Have The User Enter It!
        '-- This Can Be Changed
        szpassword = InputBox("Please Enter The Password!")

        '-- The User Did Not Enter A Password So Exit The Function
        If LenB(szpassword) = 0 Then
            Exit Function '---> Bottom
        End If

        '-- User Entered A Password So Proccess It
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

        p.ch(XX) = vbNullChar ' Put Null Terminator For C

        ZDLLPass = 0
    On Error GoTo 0

End Function

Public Function ZDLLPrnt(ByRef fname As ZipCBChar, ByVal X As Long) As Long

  '11-- Callback For Info231-zip32vc.dll - DLL Print Function

  Dim s0 As String
  Dim XX As Long

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        s0 = vbNullString

        '-- Get Info231-zip32vc.dll Message For processing
        For XX = 0 To X
            If fname.ch(XX) = 0 Then
                Exit For 'loop varying xx
              Else 'NOT FNAME.CH(XX)...
                s0 = s0 + Chr$(fname.ch(XX))
            End If
        Next XX

        '----------------------------------------------
        '-- This Is Where The DLL Passes Back Messages
        '-- To You! You Can Change The Message Printing
        '-- Below Here!
        '----------------------------------------------

        '-- Display Zip File Information
        '-- zZipInfo = zZipInfo & s0
        ''''Form1.Print s0;

        DoEvents

        ZDLLPrnt = 0
    On Error GoTo 0

End Function

Public Function ZDLLServ(ByRef mname As ZipCBChar, ByVal X As Long) As Long

  '12-- Callback For Info231-zip32vc.dll - DLL Service Function
  ' x is the size of the file

  Dim s0 As String
  Dim XX As Long

  '-- Always Put This In Callback Routines!

    On Error Resume Next

        s0 = vbNullString
        '-- Get Info231-zip32vc.dll Message For processing
        For XX = 0 To 4096
            If mname.ch(XX) = 0 Then
                Exit For 'loop varying xx
              Else 'NOT MNAME.CH(XX)...
                s0 = s0 + Chr$(mname.ch(XX))
            End If
        Next XX
        ' Form1.Print "-- " & s0 & " - " & x & " bytes"

        ' This is called for each zip entry.
        ' mname is usually the null terminated file name and x the file size.
        ' s0 has trimmed file name as VB string.

        ' At this point, s0 contains the message passed from the DLL
        ' It is up to the developer to code something useful here :)
        ZDLLServ = 0 ' Setting this to 1 will abort the zip!
    On Error GoTo 0

End Function

':)Code Fixer V3.0.9 (04/08/2005 18:02:27) 252 + 223 = 475 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 268  Code: 240  Total: 508 Lines
':) CommentOnly: 197 (38.8%)  Commented: 109 (21.5%)  Empty: 70 (13.8%)  Max Logic Depth: 3
