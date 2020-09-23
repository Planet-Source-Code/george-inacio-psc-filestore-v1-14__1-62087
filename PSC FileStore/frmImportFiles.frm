VERSION 5.00
Begin VB.Form frmImportFiles 
   Caption         =   "Import Files"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstErrorsList 
      BackColor       =   &H00FFFFC0&
      Height          =   2205
      Left            =   5850
      TabIndex        =   31
      Top             =   20788
      Width           =   4850
   End
   Begin VB.ListBox lstDuplicateList 
      BackColor       =   &H00FFFFC0&
      Height          =   2205
      Left            =   5850
      TabIndex        =   30
      Top             =   20788
      Width           =   4850
   End
   Begin VB.ListBox lstSkippedList 
      BackColor       =   &H00FFFFC0&
      Height          =   2205
      Left            =   5850
      TabIndex        =   29
      Top             =   20788
      Width           =   4850
   End
   Begin VB.ListBox lstProcessedList 
      BackColor       =   &H00FFFFC0&
      Height          =   2205
      Left            =   5850
      TabIndex        =   27
      Top             =   20788
      Width           =   4850
   End
   Begin VB.CommandButton cmdErrorsList 
      Caption         =   "Show List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3930
      TabIndex        =   26
      Top             =   2512
      Width           =   1500
   End
   Begin VB.CommandButton cmdDuplicateList 
      Caption         =   "Show List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3930
      TabIndex        =   25
      Top             =   2107
      Width           =   1500
   End
   Begin VB.CommandButton cmdSkippedList 
      Caption         =   "Show List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3930
      TabIndex        =   24
      Top             =   1704
      Width           =   1500
   End
   Begin VB.CommandButton cmdProcessedList 
      Caption         =   "Show List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3930
      TabIndex        =   23
      Top             =   1301
      Width           =   1500
   End
   Begin VB.TextBox txtFilesProcessed 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Quantity of ZIP files successfully imported."
      Top             =   1309
      Width           =   1200
   End
   Begin VB.TextBox txtLastDescription 
      Height          =   2400
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3900
      Width           =   10185
   End
   Begin VB.TextBox txtErrors 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   20
      ToolTipText     =   "Quantity of ZIP files with Errors.|The files are not ZIP files or they are corrupted."
      Top             =   2520
      Width           =   1200
   End
   Begin VB.TextBox txtDuplicateFiles 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Quantity of ZIP files duplicated.|The Title was found in the database."
      Top             =   2115
      Width           =   1200
   End
   Begin VB.ListBox lstYesNo 
      BackColor       =   &H00FFFFC0&
      Height          =   450
      ItemData        =   "frmImportFiles.frx":0000
      Left            =   6480
      List            =   "frmImportFiles.frx":000A
      TabIndex        =   15
      Top             =   21320
      Width           =   600
   End
   Begin VB.CommandButton cmdSelectFolder 
      Caption         =   "Select Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3315
      TabIndex        =   14
      ToolTipText     =   "Click Select Folder button to select the directory|with ZIP files you like to import."
      Top             =   6500
      Width           =   1600
   End
   Begin VB.TextBox txtLastTitle 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Displays the file Title current imported."
      Top             =   3210
      Width           =   10180
   End
   Begin VB.TextBox txtFilesSkipped 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Quantity of ZIP files failed to be imported.|The PSC [@PSC_ReadMe] was not found."
      Top             =   1712
      Width           =   1200
   End
   Begin VB.TextBox txtBackup 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   $"frmImportFiles.frx":0017
      Top             =   906
      Width           =   600
   End
   Begin VB.TextBox txtFilesFound 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Quantity of ZIP files found on the selected directory."
      Top             =   503
      Width           =   1200
   End
   Begin VB.TextBox txtLocationPath 
      Height          =   285
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Click Select Folder button to select the directory|with ZIP files you like to import."
      Top             =   100
      Width           =   7935
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6375
      TabIndex        =   1
      ToolTipText     =   "Returns to Main Menu."
      Top             =   6500
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Start Import"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      ToolTipText     =   "Click it to start to import ZIP files."
      Top             =   6500
      Width           =   1455
   End
   Begin VB.Label lblListName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   5850
      TabIndex        =   28
      Top             =   503
      Width           =   4850
   End
   Begin VB.Label Label9 
      Caption         =   "No of Files with Errors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   19
      Top             =   2518
      Width           =   2250
   End
   Begin VB.Label lblBackingup 
      Caption         =   "Backing up your files!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3450
      TabIndex        =   18
      Top             =   906
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label8 
      Caption         =   "No of Duplicate Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   16
      Top             =   2115
      Width           =   2250
   End
   Begin VB.Label Label7 
      Caption         =   "Last Description Processed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   3615
      Width           =   2970
   End
   Begin VB.Label Label6 
      Caption         =   "Last Title Processed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Top             =   2925
      Width           =   2250
   End
   Begin VB.Label Label5 
      Caption         =   "No of Files Skipped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   1712
      Width           =   2250
   End
   Begin VB.Label Label4 
      Caption         =   "No of Files Processed "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   1309
      Width           =   2250
   End
   Begin VB.Label Label3 
      Caption         =   "Backup Files?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   906
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "No of Files Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   503
      Width           =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Import Location Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   100
      Width           =   2250
   End
   Begin VB.Menu mnuFileItem 
      Caption         =   "File"
      Begin VB.Menu mnuExitItem 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMaintenanceItem 
      Caption         =   "Maintenance"
      Begin VB.Menu mnuEditandMovingFilesItem 
         Caption         =   "Edit and Moving Files"
      End
      Begin VB.Menu mnuCategoriesMaintItem 
         Caption         =   "Categories - Add, Edit and Delete"
      End
   End
   Begin VB.Menu mnuToolsItem 
      Caption         =   "Tools"
      Begin VB.Menu mnuScreenPositionItem 
         Caption         =   "Screen Position"
         Begin VB.Menu mnuScreenDefaultPositionItem 
            Caption         =   "Set Default Position"
         End
      End
      Begin VB.Menu mnuDatabaseItem 
         Caption         =   "Database"
         Begin VB.Menu mnuCompactRepairDatabaseItem 
            Caption         =   "Compact and repair Database"
         End
      End
      Begin VB.Menu mnuCodeDayCategoriesItem 
         Caption         =   "Code of the Day Categories"
         Begin VB.Menu mnuImportCodeDayItem 
            Caption         =   "Import Code of the Day"
         End
      End
   End
   Begin VB.Menu mnuHelpItem 
      Caption         =   "Help"
      Begin VB.Menu mnuContentsItem 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmImportFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private psDescription() As String
Private pbPscFileNotFound As Boolean
Private psNewFileName As String
Private Tooltips As New Collection
Private pbTitleExists As Boolean
Private psFileIsDuplicate As String

Private Sub Call_DoBoxBackCoulor(ByVal sBackCoulor As String)

    txtLocationPath.BackColor = sBackCoulor
    txtFilesFound.BackColor = sBackCoulor
    txtBackup.BackColor = sBackCoulor
    txtFilesProcessed.BackColor = sBackCoulor
    txtFilesSkipped.BackColor = sBackCoulor
    txtDuplicateFiles.BackColor = sBackCoulor
    txtErrors.BackColor = sBackCoulor
    txtLastTitle.BackColor = sBackCoulor
    txtLastDescription.BackColor = sBackCoulor

End Sub

Private Sub Call_DoCountBoxValues()

    If Val(txtFilesProcessed.Text) > 0 Then
        cmdProcessedList.Enabled = True
      Else 'NOT VAL(TXTFILESPROCESSED.TEXT)...
        lstProcessedList.Clear
        lstProcessedList.Top = 20788
        cmdProcessedList.Enabled = False
        lblListName.Caption = vbNullString
    End If
    If Val(txtFilesSkipped.Text) > 0 Then
        cmdSkippedList.Enabled = True
      Else 'NOT VAL(TXTFILESSKIPPED.TEXT)...
        lstSkippedList.Clear
        lstSkippedList.Top = 20788
        cmdSkippedList.Enabled = False
        lblListName.Caption = vbNullString
    End If
    If Val(txtDuplicateFiles.Text) > 0 Then
        cmdDuplicateList.Enabled = True
      Else 'NOT VAL(TXTDUPLICATEFILES.TEXT)...
        lstDuplicateList.Clear
        lstDuplicateList.Top = 20788
        cmdDuplicateList.Enabled = False
        lblListName.Caption = vbNullString
    End If
    If Val(txtErrors.Text) > 0 Then
        cmdErrorsList.Enabled = True
      Else 'NOT VAL(TXTERRORS.TEXT)...
        lstErrorsList.Clear
        lstErrorsList.Top = 20788
        cmdErrorsList.Enabled = False
        lblListName.Caption = vbNullString
    End If

End Sub

Private Sub Call_DoFileDuplicateMoveDir(ByVal sSourcePathName As String, ByVal sFullPath As String)

  Dim FSys As New FileSystemObject
  Dim lCountChars As Long
  Dim sFileName As String

    lCountChars = Len(sFullPath)
    For lCountChars = lCountChars To 1 Step -1
        If Mid$(sFullPath, lCountChars, 1) = "\" Then
            sFileName = Right$(sFullPath, Len(sFullPath) - lCountChars)
            Exit For 'loop varying lcountchars
        End If
    Next lCountChars
    If Not FSys.FolderExists(sSourcePathName & "____FilesDuplicate") Then
        FSys.CreateFolder sSourcePathName & "____FilesDuplicate"
    End If
    FSys.MoveFile sFullPath, sSourcePathName & "____FilesDuplicate\" & sFileName

End Sub

Private Sub Call_DoFileErrorMoveDir(ByVal sSourcePathName As String, ByVal sFullPath As String)

  Dim FSys As New FileSystemObject
  Dim lCountChars As Long
  Dim sFileName As String

    lCountChars = Len(sFullPath)
    For lCountChars = lCountChars To 1 Step -1
        If Mid$(sFullPath, lCountChars, 1) = "\" Then
            sFileName = Right$(sFullPath, Len(sFullPath) - lCountChars)
            Exit For 'loop varying lcountchars
        End If
    Next lCountChars
    If Not FSys.FolderExists(sSourcePathName & "____FilesWithErrors") Then
        FSys.CreateFolder sSourcePathName & "____FilesWithErrors"
    End If
    FSys.MoveFile sFullPath, sSourcePathName & "____FilesWithErrors\" & sFileName

End Sub

Private Sub Call_DoFindPscFile()

  Dim sFindFile As String

    sFindFile = Dir(App.Path & "\$$VCSTempUnzip$$\@PSC_ReadMe_*.txt")
    If Len(Trim$(sFindFile)) > 0 Then
        Call_DoReadPscFile sFindFile
      Else 'NOT LEN(TRIM$(SFINDFILE))...
        sFindFile = Dir(App.Path & "\$$VCSTempUnzip$$\File_Id.Diz")
        If Len(Trim$(sFindFile)) > 0 Then
            Call_DoReadDizFile sFindFile
          Else 'NOT LEN(TRIM$(SFINDFILE))...
            pbPscFileNotFound = True
        End If
    End If

End Sub

Private Sub Call_DoGoneOut()

  'Saving to the register the size and position of form before leaving the form

    With Me
        SaveSetting "PSC Soft", "PSC FileStore", "Height", .Height
        SaveSetting "PSC Soft", "PSC FileStore", "Left", .Left
        SaveSetting "PSC Soft", "PSC FileStore", "Top", .Top
        SaveSetting "PSC Soft", "PSC FileStore", "Width", .Width
    End With 'Me
    Set Tooltips = Nothing
    Set frmImportFiles = Nothing

End Sub

Private Sub Call_DoHideLists()

    lblListName.Caption = vbNullString
    lstProcessedList.Top = 20788
    lstSkippedList.Top = 20788
    lstDuplicateList.Top = 20788
    lstErrorsList.Top = 20788

End Sub

Private Sub Call_DoMakeBackup(ByVal sSourcePathName As String)

  Dim FSys As New FileSystemObject

    If FSys.FolderExists(sSourcePathName & "____FilesBakedUp") Then
        FSys.DeleteFolder sSourcePathName & "____FilesBakedUp"
        DoEvents
    End If
    FSys.CreateFolder sSourcePathName & "____FilesBakedUp"
    DoEvents
    FSys.CopyFile sSourcePathName & "*.*", sSourcePathName & "____FilesBakedUp"
    DoEvents

End Sub

Private Sub Call_DoMakeFilesToBeMovedDir()

  Dim FSys As New FileSystemObject

    If Not FSys.FolderExists(App.Path & "\ZipFiles\### Imported Files To Be Moved") Then
        FSys.CreateFolder (App.Path & "\ZipFiles\### Imported Files To Be Moved")
    End If

End Sub

Private Sub Call_DoMakeScreenshotDir()

  Dim FSys As New FileSystemObject

    If Not FSys.FolderExists(App.Path & "\ScreenshotPics") Then
        FSys.CreateFolder (App.Path & "\ScreenshotPics")
    End If

End Sub

Private Sub Call_DoMakeTempUnzipDir()

  Dim FSys As New FileSystemObject

    On Error Resume Next
        If FSys.FolderExists(App.Path & "\$$VCSTempUnzip$$") Then
            FSys.DeleteFolder (App.Path & "\$$VCSTempUnzip$$"), True
        End If
        FSys.CreateFolder (App.Path & "\$$VCSTempUnzip$$")
    On Error GoTo 0

End Sub

Private Sub Call_DoNoZipFiles(ByVal sNoZipFiles As String)

    MsgBox "Are NO ZIP files in the chosen directory [" & sNoZipFiles & "]!" _
           & vbNewLine & "Please select other directory with some ZIP files." _
           & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Private Sub Call_DoReadDizFile(ByVal sDizFileName As String)

  Dim FSys As New FileSystemObject
  Dim tsFile As TextStream
  Dim sDizTitle As String
  Dim sDizDescription As String

    Set tsFile = FSys.OpenTextFile(App.Path & "\$$VCSTempUnzip$$\" & sDizFileName, ForReading, False)
    sDizTitle = tsFile.ReadLine
    Do While tsFile.AtEndOfStream <> True
        sDizDescription = sDizDescription & tsFile.ReadLine
    Loop
    tsFile.Close
    sDizDescription = Func_SrchReplace(sDizTitle & sDizDescription)
    sDizTitle = Func_SrchReplace(sDizTitle)
    sDizTitle = Func_ClearChar0To31(sDizTitle)
    sDizTitle = Left$(sDizTitle, 45)
    txtLastTitle.Text = sDizTitle
    txtLastDescription.Text = sDizDescription
    Call_DoUpdateTables sDizTitle, vbNullString, sDizDescription

End Sub

Private Sub Call_DoReadPscFile(ByVal sPscFileName As String)

  Dim FSys As New FileSystemObject
  Dim tsFile As TextStream
  Dim sLineTemp As String
  Dim sPscTitle As String
  Dim sPscDescription As Variant
  Dim sPscPageAddress As Variant
  Dim lCountLines As Long
  Dim lLengh As Long
  Dim lLeft As Long
  Dim lCount As Long

    ReDim psDescription(0)
    Set tsFile = FSys.OpenTextFile(App.Path & "\$$VCSTempUnzip$$\" & sPscFileName, ForReading, False)
    'Reading @PSC_ReadMe_*.txt file
    Do While tsFile.AtEndOfStream <> True
        lCountLines = lCountLines + 1
        sLineTemp = tsFile.ReadLine
        Select Case True
          Case Left$(sLineTemp, 19) = "The author may have"
            Exit Do 'loop 
          Case Left$(sLineTemp, 21) = "You can view comments"
            lCount = 0
            lLengh = 0
            lLeft = 0
            lCount = Len(Trim$(sLineTemp))
            lLengh = lCount
            For lCount = 1 To lCount
                lLeft = lLeft + 1
                If Mid$(sLineTemp, lCount, 1) = ":" Then
                    lLeft = lLeft + 1
                    Exit For 'loop varying lcount
                End If
            Next lCount
            lLengh = lLengh - lLeft
            sPscPageAddress = Right$(Trim$(sLineTemp), lLengh)
            Exit Do 'loop 
          Case Len(Trim$(sLineTemp)) = 0
            '<STUB> Reason: 'Empty 'Case' structure used to avoid a default 'Case Else'
          Case Left$(sLineTemp, 14) = "This file came"
            '<STUB> Reason: 'Empty 'Case' structure used to avoid a default 'Case Else'
          Case Left$(sLineTemp, 6) = "Title:"
            lLengh = Len(Trim$(sLineTemp))
            lLengh = lLengh - 7
            sPscTitle = Trim$(Right$(sLineTemp, lLengh))
          Case Else
            sPscDescription = sPscDescription & " " & sLineTemp
        End Select
        DoEvents
    Loop
    tsFile.Close
    lLengh = 0
    lLengh = Len(Trim$(sPscDescription)) - 13
    If lLengh > 0 Then
        sPscDescription = Right$(Trim$(sPscDescription), lLengh)
      Else 'NOT LLENGH...
        sPscDescription = "PSC FileStore: The Author did not provide Description."
    End If
    sPscTitle = Func_SrchReplace(sPscTitle)
    sPscTitle = Func_ClearChar0To31(sPscTitle)
    sPscDescription = Func_SrchReplace(sPscDescription)
    txtLastTitle.Text = sPscTitle
    txtLastDescription.Text = sPscDescription
    Call_DoUpdateTables sPscTitle, sPscPageAddress, sPscDescription

End Sub

Private Sub Call_DoToolTips()

  Dim Tooltip   As cToolTip
  Dim Control   As Control
  Dim CollKey   As String
  Dim e         As Long

  'Code done by Ulli

    For Each Control In Controls 'cycle thru all controls
        With Control
            On Error Resume Next 'in case the control has no tooltiptext property
                CollKey = .ToolTipText 'try to access that property
                e = Err 'save error
            On Error GoTo 0
            If e = 0 Then 'the control has a tooltiptext property
                If Len(Trim$(.ToolTipText)) Then 'use that to create the custom tooltip
                    CollKey = .Name
                    On Error Resume Next 'in case control is not in an array of controls and therefore has no index property
                        CollKey = CollKey & "(" & .Index & ")"
                    On Error GoTo 0
                    Set Tooltip = New cToolTip
                    If Tooltip.Create(Control, Trim$(.ToolTipText), TTBalloonAlways, (TypeName(Control) = "TextBox"), TTIconInfo, CollKey) Then
                        Tooltips.Add Tooltip, CollKey 'to keep a reference to the current tool tip class instance (prevent it from being destroyed)
                        .ToolTipText = vbNullString 'kill tooltiptext so we don't get two tips
                    End If
                End If
            End If
        End With 'CONTROL
    Next Control

End Sub

Private Sub Call_DoUpdateTables(ByVal sPscTitle As String, ByVal sPscPageAddress As String, _
                                ByVal sPscDescription As String)

  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim sSearch As String
  Dim lCategoryNumber As Long
  Dim sScreenshot As String

    sSearch = "### Imported Files To Be Moved"
    psNewFileName = Func_FilterString(Trim$(sPscTitle))
    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    With rsCategories
        .FindFirst "CATEGORY_NAME = '" & sSearch & "'"
        If .NoMatch Then
            .AddNew
            .Fields("CATEGORY_NAME") = sSearch
            .Fields("CATEGORY_NUMBER") = .Fields("AutoNumber")
            .Fields("CATEGORY_PATH") = App.Path & "\ZipFiles\" & sSearch '### Imported Files To Be Moved"
            lCategoryNumber = .Fields("AutoNumber")
            .Update
            .MoveLast
          Else '.NOMATCH = FALSE/0
            lCategoryNumber = .Fields("CATEGORY_NUMBER")
        End If
        .Close
    End With 'RSCATEGORIES
    Set rsCategories = Nothing
    sScreenshot = Func_FindScreenshot(psNewFileName)
    If Len(Trim$(sScreenshot)) = 0 Then
        sScreenshot = "No Screenshot"
      Else 'NOT LEN(TRIM$(SSCREENSHOT))...
        sScreenshot = sScreenshot
    End If
    With rsFileDetails
        .FindFirst "TITLE = '" & Trim$(sPscTitle) & "'"
        If .NoMatch Then
            .AddNew
            .Fields("CATEGORY_NUMBER") = lCategoryNumber
            .Fields("TITLE") = Trim$(sPscTitle)
            .Fields("DESCRIPTION") = sPscDescription 'sTableDescription
            .Fields("PAGE_ADDRESS") = Trim$(sPscPageAddress)
            .Fields("FILE_NAME") = psNewFileName & ".zip"
            .Fields("SCREENSHOT") = sScreenshot
            .Update
            .MoveLast
          Else '.NOMATCH = FALSE/0
            pbTitleExists = True
            psFileIsDuplicate = Trim$(sPscTitle)
        End If
        .Close
    End With 'RSFILEDETAILS
    Set rsFileDetails = Nothing

End Sub

Private Sub Call_ThisFormSize()

    With Me
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

End Sub

Private Sub cmdDuplicateList_Click()

    Call_DoHideLists
    lblListName.Caption = "Duplicate Files"
    lstDuplicateList.Top = 788

End Sub

Private Sub cmdErrorsList_Click()

    Call_DoHideLists
    lblListName.Caption = "Files with Errors"
    lstErrorsList.Top = 788

End Sub

Private Sub cmdImport_Click()

  Dim sFileName As String   ' Walking filename variable.
  Dim sDirectoryPath As String
  Dim lNumElements As Long
  Dim sFullPath() As String
  Dim lCountElements As Long
  Dim lCountZipFiles As Long
  Dim lCountZipDirectoryFiles As Long
  Dim lCountZipFilesSkiped As Long
  Dim lCountZipFilesDuplicated As Long
  Dim lUnzipError As Long

    cmdImport.Enabled = False
    cmdSelectFolder.Enabled = False
    Screen.MousePointer = vbHourglass
    ReDim sFullPath(0) As String
    sDirectoryPath = Trim$(txtLocationPath.Text)
    sFileName = Dir(sDirectoryPath & "*.zip", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    Do While Len(sFileName) <> 0
        lNumElements = UBound(sFullPath()) + 1
        ReDim Preserve sFullPath(lNumElements)
        sFullPath(lNumElements) = sDirectoryPath & sFileName
        sFileName = Dir()  ' Get next file.
        lCountZipDirectoryFiles = lCountZipDirectoryFiles + 1
    Loop
    txtFilesFound.Text = lCountZipDirectoryFiles
    If lCountZipDirectoryFiles = 0 Then
        Call_DoBoxBackCoulor &HC0C0FF 'Light Red
        Call_DoNoZipFiles sDirectoryPath
      Else 'NOT LCOUNTZIPDIRECTORYFILES...
        'Old Backup directory must be removed to carry on with importing
        If Func_FindBackupExists(sDirectoryPath) Then
            If Not Func_BackupExists(sDirectoryPath) Then
                Screen.MousePointer = vbDefault
                cmdSelectFolder.Enabled = True
                Call_DoBoxBackCoulor &HC0C0FF 'Light Red
                Exit Sub '---> Bottom
            End If
        End If
        'Old FilesWithErrors directory must be removed to carry on with importing
        If Func_FindErrorExists(sDirectoryPath) Then
            If Not Func_ErrorExists(sDirectoryPath) Then
                Screen.MousePointer = vbDefault
                cmdSelectFolder.Enabled = True
                Call_DoBoxBackCoulor &HC0C0FF 'Light Red
                Exit Sub '---> Bottom
            End If
        End If
        'Old FilesDuplicate directory must be removed to carry on with importing
        If Func_FindDuplicateExists(sDirectoryPath) Then
            If Not Func_DuplicateExists(sDirectoryPath) Then
                Screen.MousePointer = vbDefault
                cmdSelectFolder.Enabled = True
                Call_DoBoxBackCoulor &HC0C0FF 'Light Red
                Exit Sub '---> Bottom
            End If
        End If
        Call_DoBoxBackCoulor &HC0FFC0 ' Light  Green
        If txtBackup.Text = "YES" Then
            lblBackingup.Visible = True
            DoEvents
            Call_DoMakeBackup sDirectoryPath
            lblBackingup.Visible = False
            DoEvents
        End If
        lCountElements = lNumElements
        For lCountElements = 1 To lCountElements
            DoEvents
            txtLastTitle.Text = vbNullString
            txtLastDescription.Text = vbNullString
            'Unzipping the file to be modified to a temporary directory
            Call_DoMakeTempUnzipDir
            Call_DoMakeFilesToBeMovedDir
            Call_DoMakeScreenshotDir
            uZipFileName = sFullPath(lCountElements)
            '-- Init Global Message Variables
            uZipInfo = vbNullString
            uZipNumber = 0   ' Holds The Number Of Zip Files
            '-- Select Info552-unzip32vc.dll Options - Change As Required!
            uPromptOverWrite = 0  ' 1 = Prompt To Overwrite
            uOverWriteFiles = 1   ' 1 = Always Overwrite Files
            uDisplayComment = 0   ' 1 = Display comment ONLY!!!
            '-- Change The Next Line To Do The Actual Unzip!
            uExtractList = 1       ' 1 = List Contents Of Zip 0 = Extract
            uHonorDirectories = 1  ' 1 = Honour Zip Directories
            '-- Select Filenames If Required
            '-- Or Just Select All Files
            uZipNames.uzFiles(0) = vbNullString
            uNumberFiles = 0
            '-- Select Filenames To Exclude From Processing
            ' Note UNIX convention!
            '   vbxnames.s(0) = "VBSYX/VBSYX.MID"
            '   vbxnames.s(1) = "VBSYX/VBSYX.SYX"
            '   numx = 2
            '-- Or Just Select All Files
            uExcludeNames.uzFiles(0) = vbNullString
            uNumberXFiles = 0
            '-- Change The Next 2 Lines As Required!
            '-- These Should Point To Your Directory
            'uZipFileName
            uExtractDir = App.Path & "\$$VCSTempUnzip$$" 'txtExtractRoot.Text
            If LenB(uExtractDir) Then
                uExtractList = 0 ' unzip if dir specified
            End If
            '-- Let's Go And Unzip Them!
            Call_VBUnZip32
            DoEvents
            If gbUnzipError Then
                'Found a ZIP file with errors or is not a ZIP file
                'This is NOT reliable. Must check it to see if is real corrupt
                'Calling the sub to move file and report error
                lstErrorsList.AddItem Func_RemovePath(uZipFileName)
                Call_DoFileErrorMoveDir sDirectoryPath, uZipFileName
                lUnzipError = lUnzipError + 1
                txtErrors.Text = lUnzipError
                gbUnzipError = False
                DoEvents
              Else 'GBUNZIPERROR = FALSE/0
                Call_DoFindPscFile
                If pbPscFileNotFound Then
                    'The PSC [@PSC_ReadMe] was not found.
                    'The File_Id.Diz also not found
                    lstSkippedList.AddItem Func_RemovePath(uZipFileName)
                    txtLastTitle.Text = vbNullString
                    txtLastDescription.Text = vbNullString
                    pbPscFileNotFound = False
                    lCountZipFilesSkiped = lCountZipFilesSkiped + 1
                    txtFilesSkipped.Text = lCountZipFilesSkiped
                    DoEvents
                  Else 'PBPSCFILENOTFOUND = FALSE/0
                    If pbTitleExists Then
                        'The Title was found in the database.
                        'Assumes it is a duplicate file
                        lstDuplicateList.AddItem Func_RemovePath(uZipFileName)
                        Call_DoFileDuplicateMoveDir sDirectoryPath, uZipFileName
                        txtLastTitle.Text = vbNullString
                        txtLastDescription.Text = vbNullString
                        pbTitleExists = False
                        psFileIsDuplicate = vbNullString
                        lCountZipFilesDuplicated = lCountZipFilesDuplicated + 1
                        txtDuplicateFiles.Text = lCountZipFilesDuplicated
                        DoEvents
                      Else 'PBTITLEEXISTS = FALSE/0
                        DoEvents
                        'We have all what we need so let's ZIP it back, and move the file
                        'With the proper Title to our collection under the application path
                        zTempDir = vbNullChar ' Temporary directory for use during zip process
                        zRootDir = vbNullChar
                        zJunkDir = 1 '0     ' 1 = Throw Away Path Names
                        zGrow = 1
                        zMove = 1
                        zLevel = Asc(9)  ' Compression Level (0 - 9)
                        '''START Import Web Page
                        'We going to place the Web Page (HTM) into the ZIP file
                        'If you do not want the Web Page then Comment out or remove the lines
                        'from '''START Import Web Page
                        'until the line '''END Import Web Page
                        'Also do the same thing the function  Private Function Func_FindWebPage() As String
                        If Len(Func_FindWebPage) > 0 Then
                            zArgc = 1
                            zZipFileName = uZipFileName
                            zZipFileNames.zFiles(0) = uExtractDir & "\*.htm"
                          Else 'NOT LEN(FUNC_FINDWEBPAGE)...
                            zArgc = 0
                            zZipFileName = uZipFileName
                        End If
                        '''END Import Web Page
                        Func_VBZip32
                        DoEvents
                        'Rename and move the file to be allocated to a proper directory later on
                        'Tip from Jim Jose
                        Name uZipFileName As App.Path & "\ZipFiles\### Imported Files To Be Moved\" & psNewFileName & ".zip"
                        lCountZipFiles = lCountZipFiles + 1
                        txtFilesProcessed.Text = lCountZipFiles
                        lstProcessedList.AddItem Func_RemovePath(uZipFileName)
                        DoEvents
                    End If
                End If
            End If
        Next lCountElements
        Call_DoBoxBackCoulor &HFFFFFF 'White
    End If
    cmdSelectFolder.Enabled = True
    Screen.MousePointer = vbDefault
    Call_DoCountBoxValues

End Sub

Private Sub cmdMainMenu_Click()

    Call_ThisFormSize
    frmStartMenu.Show
    Unload Me

End Sub

Private Sub cmdProcessedList_Click()

    Call_DoHideLists
    lblListName.Caption = "Files Processed"
    lstProcessedList.Top = 788

End Sub

Private Sub cmdSelectFolder_Click()

  Dim iNull As Long
  Dim lpIDList As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo
  Dim sFileName As String
  Dim lCountZipDirectoryFiles As Long

    txtLocationPath.Text = vbNullString
    txtFilesFound.Text = vbNullString
    txtFilesProcessed.Text = "0"
    txtFilesSkipped.Text = "0"
    txtDuplicateFiles.Text = "0"
    txtErrors.Text = "0"
    txtLastTitle.Text = vbNullString
    txtLastDescription.Text = vbNullString
    Call_DoBoxBackCoulor &HC0FFFF ' Light  Yellow
    Call_DoCountBoxValues
    'I got this from AllAPI Guide but can't get info on how change the Browse Folder caption
    'If you know how please let me know.
    With udtBI
        'Set the owner window
        .hWndOwner = Me.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("Import Location Path", vbNullString) 'lstrcat(szTitle, vbNullString)
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With 'UDTBI
    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
        If Right$(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If
        txtLocationPath.Text = sPath
    End If
    If sPath <> vbNullString Then
        txtFilesFound.Text = "0"
        sFileName = Dir(Trim$(txtLocationPath.Text) & "*.zip", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
        Do While Len(sFileName) <> 0
            sFileName = Dir()  ' Get next file.
            lCountZipDirectoryFiles = lCountZipDirectoryFiles + 1
            txtFilesFound.Text = lCountZipDirectoryFiles
        Loop
      Else 'NOT SPATH...
        MsgBox "The operation was cancelled!" _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    txtLocationPath.SetFocus

End Sub

Private Sub cmdSkippedList_Click()

    Call_DoHideLists
    lblListName.Caption = "Files Skipped"
    lstSkippedList.Top = 788

End Sub

Private Sub Form_Load()

    gsLocalForm = Me.Caption
    Me.Caption = gsProgName & " - " & Me.Caption & " - " & gsOwner
    With Me
        .Height = glFormHeight
        .Left = glFormLeft
        .Top = glFormTop
        .Width = glFormWidth
    End With 'ME
    Call_DoBoxBackCoulor &HC0FFFF ' Light  Yellow
    txtBackup.Text = "YES"
    Call_DoToolTips

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

Private Function Func_BackupExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject

    If vbYes = MsgBox("A previous backup directory with files in it exists!" _
       & vbNewLine & "Do you like to REMOVE the Backup Directory?" _
       & vbNewLine & "Click YES button to delete the directory." _
       & vbNewLine & "Click NO button to cancel this operation.", vbYesNo + vbCritical, gsLocalForm) Then
        If FSys.FolderExists(sSourcePathName & "____FilesBakedUp") Then
            FSys.DeleteFolder sSourcePathName & "____FilesBakedUp"
        End If
        Func_BackupExists = True
    End If

End Function

Private Function Func_DuplicateExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject

    If vbYes = MsgBox("A previous FilesDuplicate directory with files in it exists!" _
       & vbNewLine & "Do you like to REMOVE the FilesDuplicate Directory?" _
       & vbNewLine & "Click YES button to delete the directory." _
       & vbNewLine & "Click NO button to cancel this operation.", vbYesNo + vbCritical, gsLocalForm) Then
        If FSys.FolderExists(sSourcePathName & "____FilesDuplicate") Then
            FSys.DeleteFolder sSourcePathName & "____FilesDuplicate"
        End If
        Func_DuplicateExists = True
    End If

End Function

Private Function Func_ErrorExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject

    If vbYes = MsgBox("A previous FilesWithErrors directory with files in it exists!" _
       & vbNewLine & "Do you like to REMOVE the FilesWithErrors Directory?" _
       & vbNewLine & "Click YES button to delete the directory." _
       & vbNewLine & "Click NO button to cancel this operation.", vbYesNo + vbCritical, gsLocalForm) Then
        If FSys.FolderExists(sSourcePathName & "____FilesWithErrors") Then
            FSys.DeleteFolder sSourcePathName & "____FilesWithErrors"
        End If
        Func_ErrorExists = True
    End If

End Function

Private Function Func_FindBackupExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject
  Dim sFileName As String
  Dim lCountDirectoryFiles As Long

    If FSys.FolderExists(sSourcePathName & "____FilesBakedUp") Then
        sFileName = Dir(sSourcePathName & "____FilesBakedUp\*.*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
        Do While Len(sFileName) <> 0
            sFileName = Dir()  ' Get next file.
            lCountDirectoryFiles = lCountDirectoryFiles + 1
        Loop
        If lCountDirectoryFiles > 0 Then
            Func_FindBackupExists = True
        End If
    End If

End Function

Private Function Func_FindDuplicateExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject
  Dim sFileName As String
  Dim lCountDirectoryFiles As Long

    If FSys.FolderExists(sSourcePathName & "____FilesDuplicate") Then
        sFileName = Dir(sSourcePathName & "____FilesDuplicate\*.*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
        Do While Len(sFileName) <> 0
            sFileName = Dir()  ' Get next file.
            lCountDirectoryFiles = lCountDirectoryFiles + 1
        Loop
        If lCountDirectoryFiles > 0 Then
            Func_FindDuplicateExists = True
        End If
    End If

End Function

Private Function Func_FindErrorExists(ByVal sSourcePathName As String) As Boolean

  Dim FSys As New FileSystemObject
  Dim sFileName As String
  Dim lCountDirectoryFiles As Long

    If FSys.FolderExists(sSourcePathName & "____FilesWithErrors") Then
        sFileName = Dir(sSourcePathName & "____FilesWithErrors\*.*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
        Do While Len(sFileName) <> 0
            sFileName = Dir()  ' Get next file.
            lCountDirectoryFiles = lCountDirectoryFiles + 1
        Loop
        If lCountDirectoryFiles > 0 Then
            Func_FindErrorExists = True
        End If
    End If

End Function

Private Function Func_FindScreenshot(ByVal sScreenshot As String) As String

  Dim FSys As New FileSystemObject
  Dim sPicDir As String
  Dim sDirName As String
  Dim sFileName As String
  Dim lSelectLenght As Long

    lSelectLenght = 40
    If Len(sScreenshot) < lSelectLenght Then
        lSelectLenght = Len(sScreenshot)
      Else 'NOT LEN(SSCREENSHOT)...
        sScreenshot = Left$(sScreenshot, lSelectLenght)
    End If
    sDirName = Dir(Trim$(txtLocationPath.Text), vbDirectory Or vbHidden)  ' Even if hidden.
    Do While Len(sDirName) > 0
        ' Ignore the current and encompassing directories.
        If sDirName <> "." Then
            ' Check for directory with bitwise comparison.
            If sDirName <> ".." Then
                If GetAttr(Trim$(txtLocationPath.Text) & sDirName) And vbDirectory Then
                    If sScreenshot = Left$(sDirName, lSelectLenght) Then
                        sPicDir = Trim$(txtLocationPath.Text) & sDirName
                        Exit Do 'loop 
                    End If
                End If
            End If
        End If
        sDirName = Dir()  ' Get next subdirectory.
    Loop
    sFileName = Dir(sPicDir & "\pic*.*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    If FSys.FileExists(sPicDir & "\" & sFileName) Then
        If FileLen(sPicDir & "\" & sFileName) > 0 Then
            FSys.CopyFile sPicDir & "\" & sFileName, App.Path & "\ScreenshotPics\" & sFileName
            Func_FindScreenshot = sFileName
          Else 'NOT FILELEN(SPICDIR...
            Func_FindScreenshot = vbNullString
        End If
      Else 'NOT FSYS.FILEEXISTS(SPICDIR...
        sDirName = Dir(Trim$(App.Path & "\$$VCSTempUnzip$$\"), vbDirectory Or vbHidden)  ' Even if hidden.
        Do While Len(sDirName) > 0
            ' Ignore the current and encompassing directories.
            If sDirName <> "." Then
                ' Check for directory with bitwise comparison.
                If sDirName <> ".." Then
                    If GetAttr(Trim$(App.Path & "\$$VCSTempUnzip$$\") & sDirName) And vbDirectory Then
                        If sScreenshot = Left$(sDirName, lSelectLenght) Then
                            sPicDir = Trim$(App.Path & "\$$VCSTempUnzip$$\") & sDirName
                            Exit Do 'loop 
                        End If
                    End If
                End If
            End If
            sDirName = Dir()  ' Get next subdirectory.
        Loop
        sFileName = Dir(sPicDir & "\pic*.*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
        If FSys.FileExists(sPicDir & "\" & sFileName) Then
            FSys.CopyFile sPicDir & "\" & sFileName, App.Path & "\ScreenshotPics\" & sFileName
            Func_FindScreenshot = sFileName
          Else 'NOT FSYS.FILEEXISTS(SPICDIR...
            Func_FindScreenshot = vbNullString
        End If
    End If

End Function

Private Function Func_FindWebPage() As String

  Dim FSys As New FileSystemObject
  Dim sPageName As String
  Dim sFileName As String

    sPageName = Left$(psNewFileName, 40)
    sFileName = Dir(Trim$(txtLocationPath.Text) & sPageName & "*.htm", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    If Len(sFileName) > 0 Then
        If Not FSys.FileExists(App.Path & "\$$VCSTempUnzip$$\" & sFileName) Then
            FSys.CopyFile Trim$(txtLocationPath.Text) & sFileName, App.Path & "\$$VCSTempUnzip$$\" & sFileName
        End If
        Func_FindWebPage = sFileName
      Else 'NOT LEN(SFILENAME)...
        Func_FindWebPage = sFileName
    End If

End Function

Private Function Func_RemovePath(ByVal sPathToRemove As String) As String

  Dim lCountChars As Long

    lCountChars = Len(sPathToRemove)
    For lCountChars = lCountChars To 1 Step -1
        If Mid$(sPathToRemove, lCountChars, 1) = "\" Then
            Func_RemovePath = Right$(sPathToRemove, Len(sPathToRemove) - lCountChars)
            Exit For 'loop varying lcountchars
        End If
    Next lCountChars

End Function

Private Sub lstYesNo_Click()

    txtBackup.Text = lstYesNo.Text
    lstYesNo.Top = 20000

End Sub

Private Sub lstYesNo_LostFocus()

    lstYesNo.Top = 20000

End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuCategoriesMaintItem_Click()

    Call_ThisFormSize
    frmCategories.Show
    Unload Me

End Sub

Private Sub mnuCompactRepairDatabaseItem_Click()

    Call_ThisFormSize
    frmCompactDB.Show
    Unload Me

End Sub

Private Sub mnuContentsItem_Click()

    frmContents.Show vbModal

End Sub

Private Sub mnuEditandMovingFilesItem_Click()

    Call_ThisFormSize
    frmEditMoveFiles.Show
    Unload Me

End Sub

Private Sub mnuExitItem_Click()

    Call_DoGoneOut
    End

End Sub

Private Sub mnuImportCodeDayItem_Click()

    Call_ThisFormSize
    frmImportCodeDay.Show
    Unload Me

End Sub

Private Sub mnuScreenDefaultPositionItem_Click()

    frmScreenDefault.Show
    Unload Me

End Sub

Private Sub txtBackup_Click()

    With lstYesNo
        .Top = txtBackup.Top - 75
        .Left = txtBackup.Left + txtBackup.Width
        .Height = 500
    End With 'LSTYESNO

End Sub

Private Sub txtFilesFound_Change()

    If Val(txtFilesFound.Text) > 0 Then
        Call_DoBoxBackCoulor &HC0FFC0 ' Light  Green
        cmdImport.Enabled = True
      Else 'NOT VAL(TXTFILESFOUND.TEXT)...
        cmdImport.Enabled = False
        Call_DoBoxBackCoulor &HC0C0FF 'Light Red
    End If

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:26) 7 + 874 = 881 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-12 21:11)  Decl: 7  Code: 997  Total: 1004 Lines
':) CommentOnly: 55 (5.5%)  Commented: 73 (7.3%)  Empty: 166 (16.5%)  Max Logic Depth: 7
