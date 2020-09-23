VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStartMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnzip 
      Caption         =   "Unzip It"
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
      Left            =   8524
      TabIndex        =   21
      ToolTipText     =   $"frmStartMenu.frx":0000
      Top             =   6500
      Width           =   1960
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
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
      Left            =   6552
      TabIndex        =   20
      ToolTipText     =   $"frmStartMenu.frx":008D
      Top             =   6500
      Width           =   1960
   End
   Begin VB.Frame frameCategorylist 
      Caption         =   "Category List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4950
      Left            =   1597
      TabIndex        =   18
      Top             =   21027
      Width           =   7950
      Begin VB.ListBox lstCategoryList 
         BackColor       =   &H00FFFFC0&
         Height          =   4545
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame frameShot 
      Caption         =   "Screen Shot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6255
      Left            =   105
      TabIndex        =   13
      Top             =   20120
      Width           =   10935
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         LargeChange     =   512
         Left            =   50
         SmallChange     =   16
         TabIndex        =   16
         Top             =   5950
         Width           =   10575
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   5775
         LargeChange     =   512
         Left            =   10640
         SmallChange     =   16
         TabIndex        =   15
         Top             =   200
         Width           =   285
      End
      Begin VB.PictureBox picShot1 
         Height          =   5775
         Left            =   50
         ScaleHeight     =   5715
         ScaleWidth      =   10515
         TabIndex        =   14
         Top             =   200
         Width           =   10575
         Begin VB.PictureBox picShot2 
            AutoSize        =   -1  'True
            Height          =   5175
            Left            =   240
            ScaleHeight     =   5115
            ScaleWidth      =   10035
            TabIndex        =   17
            Top             =   240
            Width           =   10095
         End
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   5610
      TabIndex        =   11
      ToolTipText     =   $"frmStartMenu.frx":0166
      Top             =   6050
      Width           =   5385
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1710
      Width           =   5385
   End
   Begin VB.TextBox txtCategory 
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1020
      Width           =   5385
   End
   Begin VB.CommandButton cmdWebPage 
      Caption         =   "Author Web Page at PSC "
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
      Left            =   5616
      TabIndex        =   5
      ToolTipText     =   "Click the button call the Author’s Web Page at Planet Source Code.|Note: You have to be connected to the Internet."
      Top             =   240
      Width           =   5392
   End
   Begin VB.CommandButton cmdShowShot 
      Caption         =   "Show Sreenshot"
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
      Left            =   4588
      TabIndex        =   4
      ToolTipText     =   "Click to display the Screenshot."
      Top             =   6500
      Width           =   1960
   End
   Begin VB.TextBox txtDescription 
      Height          =   3225
      Left            =   5616
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5392
   End
   Begin VB.CommandButton cmdShowCategories 
      Caption         =   "Show Categories"
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
      Left            =   2623
      TabIndex        =   2
      ToolTipText     =   "If clicked the Category list will be showed.|Select the Category you like to display."
      Top             =   6500
      Width           =   1960
   End
   Begin MSComctlLib.ListView listviewTitles 
      Height          =   6100
      Left            =   136
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Select an item from the list to show the file details."
      Top             =   240
      Width           =   5392
      _ExtentX        =   9499
      _ExtentY        =   10769
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   660
      TabIndex        =   0
      ToolTipText     =   "Click the button to Close the program."
      Top             =   6500
      Width           =   1960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   12
      ImageHeight     =   15
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartMenu.frx":023F
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartMenu.frx":0745
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      Caption         =   "Search String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5610
      TabIndex        =   12
      Top             =   5740
      Width           =   5385
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5610
      TabIndex        =   10
      Top             =   2115
      Width           =   5385
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5610
      TabIndex        =   8
      Top             =   1425
      Width           =   5385
   End
   Begin VB.Label lblCategory 
      Alignment       =   2  'Center
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5610
      TabIndex        =   6
      Top             =   735
      Width           =   5385
   End
   Begin VB.Menu mnuFileItem 
      Caption         =   "File"
      Begin VB.Menu mnuImportZipFilesItem 
         Caption         =   "Import Zip Files"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
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
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tooltips As New Collection
Private psPageAddress As String
Private psFilePath As String
Private psFileName As String
Private psScreenshot As String
Private bSearchCategory As Boolean

Private Sub Call_DoClearBoxes()

    txtCategory.Text = vbNullString
    txtTitle.Text = vbNullString
    txtDescription.Text = vbNullString

End Sub

Private Sub Call_DoClearColumnHeaders()

    listviewTitles.ColumnHeaders.Clear
    listviewTitles.ColumnHeaders.Add , , "Titles", 12000

End Sub

Public Sub Call_DoDatabaseEmpty()

    MsgBox "Database is Empty!" _
           & vbNewLine & "Please enter or import files." _
           & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbCritical, gsLocalForm

End Sub

Private Sub Call_DoDeleteTempUnzipDir()

  Dim FSys As New FileSystemObject
  Dim sDirName As String

    sDirName = App.Path & "\$$VCSTempUnzip$$"
    On Error Resume Next
        If FSys.FolderExists(sDirName) Then
            FSys.DeleteFolder (sDirName), True
        End If
    On Error GoTo 0

End Sub

Private Sub Call_DoEnableButton()

    cmdShowCategories.Enabled = True
    cmdSearch.Enabled = Len(Trim$(txtSearch.Text)) > 0
    cmdUnzip.Enabled = Len(Trim$(txtTitle.Text)) > 0

End Sub

Private Sub Call_DoFindCategory(ByVal sCategoryToShow As String)

  Dim lvListItem As ListItem
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim lCountRecords As Long
  Dim lCategoryNumber As Long
  Dim lTitlesFound As Long
  Dim sIfGreaterThenZero As String

    Call_DoClearColumnHeaders
    Screen.MousePointer = vbHourglass
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    With rsCategories
        .FindFirst "CATEGORY_NAME = '" & sCategoryToShow & "'"
        If Not .NoMatch Then
            lCategoryNumber = .Fields("CATEGORY_NUMBER")
        End If
        .Close
    End With 'RSCATEGORIES
    If lCategoryNumber > 0 Then
        sIfGreaterThenZero = " Category"
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails WHERE CATEGORY_NUMBER = " & lCategoryNumber)
      Else 'NOT LCATEGORYNUMBER...
        sIfGreaterThenZero = vbNullString
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    End If
    If rsFileDetails.RecordCount > 0 Then
        With rsFileDetails
            .MoveFirst
            .MoveLast
            lCountRecords = .RecordCount
            .MoveFirst
            For lCountRecords = 1 To lCountRecords
                Set lvListItem = listviewTitles.ListItems.Add(, , .Fields("TITLE"))
                lTitlesFound = lTitlesFound + 1
                .MoveNext
            Next lCountRecords
            .Close
        End With 'RSFILEDETAILS
      Else 'NOT RSFILEDETAILS.RECORDCOUNT...
        Call_DoCategoryNoFiles sCategoryToShow
    End If
    Screen.MousePointer = vbDefault
    listviewTitles.ColumnHeaders.Clear
    listviewTitles.ColumnHeaders.Add , , lTitlesFound & " Titles found on " & sCategoryToShow & sIfGreaterThenZero, 12000
    Set rsFileDetails = Nothing

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
    Set frmStartMenu = Nothing

End Sub

Private Sub Call_DoListCategory()

  Dim rsCategories As Recordset
  Dim lCountRecords As Long

    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories ORDER By CATEGORY_NAME")
    lstCategoryList.Clear
    frameCategorylist.Top = 1027
    frameCategorylist.Caption = "Categories List"
    With rsCategories
        .MoveFirst
        .MoveLast
        lCountRecords = .RecordCount
        .MoveFirst
        lstCategoryList.AddItem "All Categories"
        For lCountRecords = 1 To lCountRecords
            lstCategoryList.AddItem .Fields("CATEGORY_NAME")
            .MoveNext
        Next lCountRecords
        .Close
    End With 'RSCATEGORIES
    Set rsCategories = Nothing
    lstCategoryList.SetFocus

End Sub

Private Sub Call_DoMakeMainCategoryDir()

  Dim FSys As New FileSystemObject
  Dim rsCategories As Recordset
  Dim sDirectoryName As String
  Dim sCategoryPath As String

  'If it is Categories in the table and doesn't have directories on disk the directories
  'will be created and the path will be placed on the table

    If Not FSys.FolderExists(App.Path & "\ZipFiles") Then
        FSys.CreateFolder App.Path & "\ZipFiles"
    End If
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    If rsCategories.RecordCount > 0 Then
        With rsCategories
            .MoveFirst
            Do While Not .EOF
                sDirectoryName = Func_FilterString(.Fields("CATEGORY_NAME"))
                sCategoryPath = App.Path & "\ZipFiles\" & sDirectoryName
                .Edit
                .Fields("CATEGORY_PATH") = sCategoryPath
                .Update
                If Not FSys.FolderExists(sCategoryPath) Then
                    FSys.CreateFolder sCategoryPath
                End If
                .MoveNext
            Loop
            .Close
        End With 'RSCATEGORIES
    End If
    Set rsCategories = Nothing

End Sub

Private Sub Call_DoSearchCategory(ByVal sCategoryToShow As String)

  Dim lvListItem As ListItem
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim lCountRecords As Long
  Dim lCategoryNumber As Long
  Dim lTitlesFound As Long
  Dim sIfGreaterThenZero As String

    Call_DoClearColumnHeaders
    Screen.MousePointer = vbHourglass
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    With rsCategories
        .FindFirst "CATEGORY_NAME = '" & sCategoryToShow & "'"
        If Not .NoMatch Then
            lCategoryNumber = .Fields("CATEGORY_NUMBER")
        End If
        .Close
    End With 'RSCATEGORIES
    If lCategoryNumber > 0 Then
        sIfGreaterThenZero = " Category"
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails WHERE CATEGORY_NUMBER = " & _
                            lCategoryNumber & " AND (Title LIKE '*" & Trim$(txtSearch) & "*' OR Description LIKE '*" & Trim$(txtSearch) & "*');")
      Else 'NOT LCATEGORYNUMBER...
        sIfGreaterThenZero = vbNullString
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails" & " WHERE Title LIKE '*" & _
                            Trim$(txtSearch) & "*' OR Description LIKE '*" & Trim$(txtSearch) & "*';")
    End If
    If rsFileDetails.RecordCount > 0 Then
        With rsFileDetails
            .MoveFirst
            .MoveLast
            lCountRecords = .RecordCount
            .MoveFirst
            For lCountRecords = 1 To lCountRecords
                Set lvListItem = listviewTitles.ListItems.Add(, , .Fields("TITLE"))
                lTitlesFound = lTitlesFound + 1
                .MoveNext
            Next lCountRecords
            .Close
        End With 'RSFILEDETAILS
      Else 'NOT RSFILEDETAILS.RECORDCOUNT...
        Call_DoCategoryNoFiles sCategoryToShow
    End If
    Screen.MousePointer = vbDefault
    listviewTitles.ColumnHeaders.Clear
    listviewTitles.ColumnHeaders.Add , , lTitlesFound & " Titles found on " & sCategoryToShow & sIfGreaterThenZero, 12000
    Set rsFileDetails = Nothing

End Sub

Private Sub Call_DoSetSreenReg()

    With Me
        .Height = 7815
        .Width = 11265
        .Move ((Screen.Width - .Width) \ 2), ((Screen.Height - .Height) \ 2)
        SaveSetting "PSC Soft", "PSC FileStore", "Height", .Height
        SaveSetting "PSC Soft", "PSC FileStore", "Left", .Left
        SaveSetting "PSC Soft", "PSC FileStore", "Top", .Top
        SaveSetting "PSC Soft", "PSC FileStore", "Width", .Width
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

End Sub

Private Sub Call_DoShowPicture(ByVal sNewPicture As String)

    picShot2.Move 0, 0
    'Lee Weiner - 03/08/99
    'Move the display box to the upper-left corner of the
    'container
    picShot2.Picture = LoadPicture(App.Path & "\ScreenshotPics\" & sNewPicture)
    'If the new width of the display is less than the width of
    'the container, disable the horizontal scroll bar.  If not,
    'set the Max of the scroll bar to the difference in width.
    VScroll1.Enabled = True
    HScroll1.Enabled = True
    If picShot2.Width <= picShot1.Width Then
        HScroll1.Enabled = False
      Else 'NOT picShot2.WIDTH...
        HScroll1.Max = picShot2.Width - picShot1.ScaleWidth
        HScroll1.Value = 0
    End If
    'If the new height of the display is less than the height of
    'the container, disable the vertical scroll bar.  If not,
    'set the Max of the scroll bar to the difference in height.
    If picShot2.Height <= picShot1.Height Then
        VScroll1.Enabled = False
      Else 'NOT picShot2.HEIGHT...
        VScroll1.Max = picShot2.Height - picShot1.ScaleHeight
        VScroll1.Value = 0
    End If

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

Private Sub Call_ThisFormSize()

    With Me
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

End Sub

Private Sub cmdExit_Click()

    Call_DoGoneOut
    End

End Sub

Private Sub cmdSearch_Click()

    bSearchCategory = True
    cmdShowCategories.Enabled = False
    cmdSearch.Enabled = False
    cmdShowShot.Enabled = False
    cmdUnzip.Enabled = False
    cmdWebPage.Enabled = False
    Call_DoClearBoxes
    listviewTitles.ListItems.Clear
    Call_DoListCategory

End Sub

Private Sub cmdShowCategories_Click()

    bSearchCategory = False
    cmdShowCategories.Enabled = False
    cmdSearch.Enabled = False
    cmdShowShot.Enabled = False
    cmdUnzip.Enabled = False
    cmdWebPage.Enabled = False
    Call_DoClearBoxes
    listviewTitles.ListItems.Clear

    Call_DoListCategory

End Sub

Private Sub cmdShowShot_Click()

    Select Case True
      Case frameShot.Top < 20000
        Call_DoEnableButton
        cmdShowShot.Caption = "Show Shot"
        frameShot.Top = 20120
        listviewTitles.Visible = True
        cmdWebPage.Visible = True
        lblCategory.Visible = True
        txtCategory.Visible = True
        lblTitle.Visible = True
        txtTitle.Visible = True
        lblDescription.Visible = True
        txtDescription.Visible = True
        lblSearch.Visible = True
        txtSearch.Visible = True
        listviewTitles.SetFocus
      Case frameShot.Top > 20000
        cmdShowCategories.Enabled = False
        cmdSearch.Enabled = False
        cmdUnzip.Enabled = False
        cmdShowShot.Caption = "Hide Shot"
        listviewTitles.Visible = False
        cmdWebPage.Visible = False
        lblCategory.Visible = False
        txtCategory.Visible = False
        lblTitle.Visible = False
        txtTitle.Visible = False
        lblDescription.Visible = False
        txtDescription.Visible = False
        lblSearch.Visible = False
        txtSearch.Visible = False
        frameShot.Top = 120
        Call_DoShowPicture psScreenshot
    End Select

End Sub

Private Sub cmdUnzip_Click()

  Dim FSys As New FileSystemObject
  Dim iNull As Long
  Dim lpIDList As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo
  Dim sTargetPath As String
  Dim sDirName As String

  'I got this from AllAPI Guide but can't get info on how change the Browse Folder caption
  'If you know how please let me know.

    sDirName = Left$(psFileName, (Len(psFileName) - 4))
    With udtBI
        'Set the owner window
        .hWndOwner = Me.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("Unzip Location Path", vbNullString) 'lstrcat(szTitle, vbNullString)
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
        sTargetPath = sPath
    End If
    If sTargetPath <> vbNullString Then
        If Left$(sTargetPath, Len(App.Path)) <> App.Path Then
            With FSys
                If .FolderExists(sTargetPath & sDirName) Then
                    .DeleteFolder (sTargetPath & sDirName)
                    .CreateFolder (sTargetPath & sDirName)
                  Else 'NOT .FOLDEREXISTS(STARGETPATH...
                    .CreateFolder (sTargetPath & sDirName)
                End If
            End With 'FSYS
            uZipFileName = psFilePath & "\" & sDirName
            '-- Init Global Message Variables
            uZipInfo = vbNullString
            uZipNumber = 0   ' Holds The Number Of Zip Files
            '-- Select Info552-unzip32vc.dll Options - Change As Required!
            uPromptOverWrite = 1  ' 1 = Prompt To Overwrite
            uOverWriteFiles = 0   ' 1 = Always Overwrite Files
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
            uExtractDir = sTargetPath & sDirName
            If LenB(uExtractDir) Then
                uExtractList = 0 ' unzip if dir specified
            End If
            '-- Let's Go And Unzip Them!
            Call_VBUnZip32
          Else 'NOT STARGETPATH... 'NOT LEFT$(STARGETPATH,...
            MsgBox "Unzip directory selected  [" & sTargetPath & " ]" _
                   & vbNewLine & "The Unzip directory CANNOT be the application directory." _
                   & vbNewLine & "Please select other Unzip directory." _
                   & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbCritical, gsLocalForm
        End If
      Else 'NOT STARGETPATH...
        MsgBox "The operation was cancelled!" _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    listviewTitles.SetFocus

End Sub

Private Sub cmdWebPage_Click()

    ShellExecute Me.hwnd, "open", psPageAddress, vbNullString, "C:\", 5

End Sub

Private Sub Form_Load()

  Dim rsFileDetails As Recordset

    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    gsLocalForm = Me.Caption
    Me.Caption = gsProgName & " - " & Me.Caption & " - " & gsOwner
    Select Case True
      Case Me.Height < 4000
        Call_DoSetSreenReg
      Case Me.Width < 6000
        Call_DoSetSreenReg
      Case (GetSetting("PSC Soft", "PSC FileStore", "Height")) = vbNullString
        Call_DoSetSreenReg
      Case (GetSetting("PSC Soft", "PSC FileStore", "Left")) = vbNullString
        Call_DoSetSreenReg
      Case (GetSetting("PSC Soft", "PSC FileStore", "Top")) = vbNullString
        Call_DoSetSreenReg
      Case (GetSetting("PSC Soft", "PSC FileStore", "Width")) = vbNullString
        Call_DoSetSreenReg
    End Select
    If Not gbFormT Then
        glFormHeight = GetSetting("PSC Soft", "PSC FileStore", "Height")
        glFormLeft = GetSetting("PSC Soft", "PSC FileStore", "Left")
        glFormTop = GetSetting("PSC Soft", "PSC FileStore", "Top")
        glFormWidth = GetSetting("PSC Soft", "PSC FileStore", "Width")
        gbFormT = True
    End If
    Select Case True
      Case glFormHeight < 4000
        Call_DoSetSreenReg
      Case glFormWidth < 6000
        Call_DoSetSreenReg
    End Select
    With Me
        .Height = glFormHeight
        .Left = glFormLeft
        .Top = glFormTop
        .Width = glFormWidth
    End With 'ME
    'Make Categories Path and Directories
    'Just in case the Category Names have no Category Path and Directories
    Call_DoMakeMainCategoryDir
    Call_DoDeleteTempUnzipDir
    listviewTitles.ColumnHeaders.Add , , "Titles", 12000
    If rsFileDetails.RecordCount < 1 Then
        cmdShowCategories.Enabled = False
        txtSearch.Locked = True
        Call_DoDatabaseEmpty
    End If
    Call_DoToolTips
    rsFileDetails.Close
    Set rsFileDetails = Nothing
    cmdShowShot.Enabled = False
    cmdSearch.Enabled = False
    cmdUnzip.Enabled = False
    VScroll1.Max = picShot2.Height - picShot1.ScaleHeight

End Sub

Private Sub Form_Paint()

    VScroll1.Max = picShot2.Height - picShot1.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

Private Sub HScroll1_Change()

  'Move the display picturebox laterally in response to changes
  'in the horizontal scroll bar.

    picShot2.Left = -1 * HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()

    picShot2.Left = -1 * HScroll1.Value

End Sub

Private Sub listviewTitles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  Dim lSortOrder As Long
  Dim lngIndex As Long

    For lngIndex = 1 To listviewTitles.ColumnHeaders.Count - 1
        'Loop through all of the column headers
        'And dectroy it's icon
        listviewTitles.ColumnHeaders(lngIndex).Icon = 0
    Next lngIndex
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    listviewTitles.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lSortOrder = listviewTitles.SortOrder
    Select Case lSortOrder
      Case 0
        listviewTitles.SortOrder = lvwDescending
        listviewTitles.ColumnHeaders(ColumnHeader.Index).Icon = 2
      Case 1
        listviewTitles.SortOrder = lvwAscending
        listviewTitles.ColumnHeaders(ColumnHeader.Index).Icon = 1
    End Select
    listviewTitles.Sorted = True

End Sub

Private Sub listviewTitles_ItemClick(ByVal Item As MSComctlLib.ListItem)

  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset

    psScreenshot = vbNullString
    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    With rsFileDetails
        .FindFirst "TITLE = '" & listviewTitles.SelectedItem & "'"
        If Not .NoMatch Then
            rsCategories.FindFirst "CATEGORY_NUMBER = " & .Fields("CATEGORY_NUMBER")
            If Not rsCategories.NoMatch Then
                txtCategory = rsCategories.Fields("CATEGORY_NAME")
            End If
            txtTitle.Text = Trim$(.Fields("TITLE"))
            txtDescription = Trim$(.Fields("DESCRIPTION"))
            psPageAddress = Trim$(.Fields("PAGE_ADDRESS"))
            'Placing the File Name and Path in memory in case the user wants to Unzip it
            psFilePath = Trim$(rsCategories.Fields("CATEGORY_PATH"))
            psFileName = Trim$(.Fields("FILE_NAME"))
            psScreenshot = Trim$(.Fields("SCREENSHOT"))
            cmdShowShot.Enabled = Not (psScreenshot = "No Screenshot")
        End If
        rsCategories.Close
        .Close
        cmdWebPage.Enabled = Left$(psPageAddress, 7) = "http://"
    End With 'RSFILEDETAILS
    Set rsFileDetails = Nothing
    Set rsCategories = Nothing

End Sub

Private Sub lstCategoryList_Click()

    If bSearchCategory Then
        Call_DoSearchCategory lstCategoryList.Text
      Else 'BSEARCHCATEGORY = FALSE/0
        Call_DoFindCategory lstCategoryList.Text
    End If
    Call_DoEnableButton
    frameCategorylist.Top = 21027
    lstCategoryList.Clear

End Sub

Private Sub lstCategoryList_LostFocus()

    Call_DoClearColumnHeaders
    Call_DoEnableButton
    frameCategorylist.Top = 21027
    lstCategoryList.Clear

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

Private Sub mnuImportZipFilesItem_Click()

    Call_ThisFormSize
    frmImportFiles.Show
    Unload Me

End Sub

Private Sub mnuScreenDefaultPositionItem_Click()

    frmScreenDefault.Show
    Unload Me

End Sub

Private Sub txtSearch_Change()

    cmdSearch.Enabled = Len(Trim$(txtSearch.Text)) > 0

End Sub

Private Sub txtTitle_Change()

    cmdUnzip.Enabled = Len(Trim$(txtTitle.Text)) > 0

End Sub

Private Sub VScroll1_Change()

  'Move the display picturebox vertically in response to
  'changes in the vertical scroll bar.

    picShot2.Top = -1 * VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()

    picShot2.Top = -1 * VScroll1.Value

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:22) 7 + 750 = 757 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 7  Code: 759  Total: 766 Lines
':) CommentOnly: 50 (6.5%)  Commented: 41 (5.4%)  Empty: 144 (18.8%)  Max Logic Depth: 6
