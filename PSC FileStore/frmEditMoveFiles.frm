VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditMoveFiles 
   Caption         =   "Edit and Moving Files"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleMode       =   0  'User
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCategorize 
      Caption         =   "Categorize"
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
      Left            =   9420
      TabIndex        =   26
      ToolTipText     =   $"frmEditMoveFiles.frx":0000
      Top             =   5850
      Width           =   1500
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
      Left            =   1300
      TabIndex        =   12
      Top             =   21002
      Width           =   7950
      Begin VB.ListBox lstCategoryList 
         BackColor       =   &H00FFFFC0&
         Height          =   4545
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7695
      End
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
      Left            =   9420
      TabIndex        =   7
      ToolTipText     =   "Returns to Main Menu."
      Top             =   6225
      Width           =   1500
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
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
      Left            =   9420
      TabIndex        =   1
      ToolTipText     =   "Click the button to select from the|list the Category to be loaded."
      Top             =   3600
      Width           =   1500
   End
   Begin VB.CommandButton cmdMoveFiles 
      Caption         =   "Move"
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
      Left            =   9420
      TabIndex        =   2
      ToolTipText     =   "Click the button to select from the list the|Category name you would like to move the files."
      Top             =   3975
      Width           =   1500
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   9420
      TabIndex        =   3
      ToolTipText     =   "By clicking this button the boxes becomes enable to be Edit.|After the changes are done click the Save button."
      Top             =   4350
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   9420
      TabIndex        =   4
      ToolTipText     =   "Save the changes that have been made."
      Top             =   4725
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   9420
      TabIndex        =   5
      ToolTipText     =   "Stop the current operation."
      Top             =   5100
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   9420
      TabIndex        =   6
      ToolTipText     =   $"frmEditMoveFiles.frx":0173
      Top             =   5475
      Width           =   1500
   End
   Begin VB.Frame frameTitles 
      Caption         =   "Titles Names"
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
      Height          =   2655
      Left            =   225
      TabIndex        =   11
      Top             =   3975
      Width           =   9000
      Begin VB.ListBox lstTitles 
         Height          =   2205
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click an entry to view the record."
         Top             =   270
         Width           =   8750
      End
   End
   Begin VB.OptionButton optOneAll 
      Caption         =   "Move All Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      ToolTipText     =   $"frmEditMoveFiles.frx":0203
      Top             =   3600
      Width           =   2250
   End
   Begin VB.OptionButton optOneAll 
      Caption         =   "Move One File Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   0
      Left            =   2070
      TabIndex        =   9
      ToolTipText     =   "Click to select this Check Box if you want to move|only one file to a different Category."
      Top             =   3600
      Value           =   -1  'True
      Width           =   2250
   End
   Begin VB.Frame frameRecResults 
      Caption         =   "Current Record Results"
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
      Height          =   3375
      Left            =   225
      TabIndex        =   8
      Top             =   100
      Width           =   10695
      Begin VB.PictureBox picCFXPBugFixfrmEditMove 
         BorderStyle     =   0  'None
         Height          =   3038
         Left            =   100
         ScaleHeight     =   3045
         ScaleWidth      =   10500
         TabIndex        =   14
         Top             =   276
         Width           =   10495
         Begin VB.CommandButton cmdGetShot 
            Caption         =   "Get Shot"
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
            Left            =   8735
            TabIndex        =   25
            ToolTipText     =   " Click to get a Screenshot picture file from the disk."
            Top             =   1017
            Width           =   1500
         End
         Begin VB.TextBox txtShot 
            Height          =   285
            Left            =   1440
            TabIndex        =   24
            ToolTipText     =   $"frmEditMoveFiles.frx":02B8
            Top             =   1017
            Width           =   7250
         End
         Begin VB.TextBox txtWebAddress 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "A Author's specific Web Page located on Planet Source Code."
            Top             =   672
            Width           =   8795
         End
         Begin VB.TextBox txtDescription 
            Height          =   1335
            Left            =   140
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Describes what the program is and what can it do."
            Top             =   1647
            Width           =   10095
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   1440
            TabIndex        =   16
            ToolTipText     =   "Code brief description."
            Top             =   327
            Width           =   8795
         End
         Begin VB.TextBox txtCategory 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            ToolTipText     =   "The Category name that was allocated at|Maintenance time or at Import time."
            Top             =   -18
            Width           =   8795
         End
         Begin VB.Label Label5 
            Caption         =   "Screenshot"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   140
            TabIndex        =   23
            Top             =   1017
            Width           =   1300
         End
         Begin VB.Label Label4 
            Caption         =   "Web Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   140
            TabIndex        =   21
            Top             =   672
            Width           =   1300
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   140
            TabIndex        =   20
            Top             =   1362
            Width           =   1300
         End
         Begin VB.Label Label2 
            Caption         =   "Title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   140
            TabIndex        =   19
            Top             =   327
            Width           =   1300
         End
         Begin VB.Label Label1 
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   140
            TabIndex        =   18
            Top             =   -18
            Width           =   1300
         End
      End
   End
   Begin MSComDlg.CommonDialog cdGetShot 
      Left            =   10680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitlesFound 
      Caption         =   "0 Titles Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6570
      TabIndex        =   27
      Top             =   3600
      Width           =   2505
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
Attribute VB_Name = "frmEditMoveFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tooltips As New Collection
Private pbAllCategories As Boolean
Private psButtonSelect As String
Private psCategoryName As String
Private psOldTitle As String
Private psOldAddress As String
Private psOldShot As String
Private psCurrentCategory As String
Private pbEdit As Boolean
Private psNewShotFileFromPath As String

Private Sub Call_DoBackCoulor(ByVal sBackColour As String)

    txtCategory.BackColor = sBackColour
    txtTitle.BackColor = sBackColour
    txtDescription.BackColor = sBackColour
    txtWebAddress.BackColor = sBackColour
    txtShot.BackColor = sBackColour

End Sub

Private Sub Call_DoBoxClear()

    txtCategory.Text = vbNullString
    txtTitle.Text = vbNullString
    txtDescription.Text = vbNullString
    txtWebAddress.Text = vbNullString
    txtShot.Text = vbNullString

End Sub

Private Sub Call_DoClearText()

  'Call Function to replace Single Quote with one Apostrophe and Double Quotes with two Apostrophes

    txtTitle.Text = Func_SrchReplace(txtTitle.Text)
    txtDescription.Text = Func_SrchReplace(txtDescription.Text)
    txtWebAddress.Text = Func_SrchReplace(txtWebAddress.Text)

End Sub

Private Sub Call_DoCodeDayDelTitle(ByVal sDelTitle As String)

  Dim rsCodeDay As Recordset

    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    With rsCodeDay
        .FindFirst "CODE_TITLE = '" & sDelTitle & "'"
        If Not .NoMatch Then
            .Delete
        End If
        .Close
    End With 'RSCODEDAY
    Set rsCodeDay = Nothing

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
    Set frmEditMoveFiles = Nothing

End Sub

Private Sub Call_DoListCategory()

  Dim rsCategories As Recordset
  Dim lCountRecords As Long

    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories ORDER BY CATEGORY_NAME")
    If psButtonSelect = "Load" Then
        Call_DoBoxClear
        lstTitles.Clear
        lstCategoryList.Clear
        lstCategoryList.AddItem "All Categories"
    End If
    If rsCategories.RecordCount > 0 Then
        With rsCategories
            .MoveFirst
            .MoveLast
            lCountRecords = .RecordCount
            .MoveFirst
            For lCountRecords = 1 To lCountRecords
                If psButtonSelect = "MoveFiles" Then
                    If psCategoryName <> .Fields("CATEGORY_NAME") Then
                        lstCategoryList.AddItem .Fields("CATEGORY_NAME")
                    End If
                  Else 'NOT PSBUTTONSELECT...
                    lstCategoryList.AddItem .Fields("CATEGORY_NAME")
                End If
                .MoveNext
            Next lCountRecords
            .Close
            If psButtonSelect = "MoveFiles" Then
                If lstCategoryList.ListCount <= 0 Then
                    frameCategorylist.Top = 22000
                    MsgBox "No Categories found to move Tile!" _
                           & vbNewLine & "Please Create some Category names." _
                           & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
                End If
            End If
        End With 'RSCATEGORIES
      Else 'NOT RSCATEGORIES.RECORDCOUNT...
        rsCategories.Close
        cmdLoad.Enabled = False
        cmdCategorize.Enabled = False
        frameCategorylist.Top = 22000
        MsgBox "No Categories found on table!" _
               & vbNewLine & "Please Create some Category names." _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    Set rsCategories = Nothing
    lstCategoryList.SetFocus

End Sub

Private Sub Call_DoListTitles(ByVal lCategoryNumber As Long, ByVal sListCategoryName As String)

  Dim rsFileDetails As Recordset
  Dim lCountRecords As Long
  Dim lCountFoundTitles As Long

    lstTitles.Clear
    lblTitlesFound.Caption = "0 Titles Found"
    If lCategoryNumber > 0 Then
        pbAllCategories = False
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails WHERE CATEGORY_NUMBER = " & lCategoryNumber)
      Else 'NOT LCATEGORYNUMBER...
        pbAllCategories = True
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    End If
    If rsFileDetails.RecordCount > 0 Then
        Call_DoBackCoulor &HFFFFFF 'White
        lstTitles.Enabled = True
        lstTitles.BackColor = &HFFFFFF 'White
        With rsFileDetails
            .MoveFirst
            .MoveLast
            lCountRecords = .RecordCount
            .MoveFirst
            For lCountRecords = 1 To lCountRecords
                lstTitles.AddItem .Fields("TITLE")
                lblTitlesFound.Caption = "0 Titles Found"
                lCountFoundTitles = lCountFoundTitles + 1
                lblTitlesFound.Caption = lCountFoundTitles & " Titles Found"
                .MoveNext
            Next lCountRecords
            .Close
        End With 'RSFILEDETAILS
      Else 'NOT RSFILEDETAILS.RECORDCOUNT...
        Call_DoBoxClear
        Call_DoLockAll
        Call_DoCategoryNoFiles sListCategoryName
    End If
    Set rsFileDetails = Nothing

End Sub

Private Sub Call_DoLockAll()

    Call_DoLockBoxes
    Call_DoBackCoulor &H8000000F  'Grey
    With lstTitles
        .Clear
        .Enabled = False
        .BackColor = &H8000000F  'Grey
    End With 'lstTitles
    optOneAll(0).Enabled = False
    optOneAll(1).Enabled = False
    cmdMoveFiles.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub Call_DoLockBoxes()

    txtCategory.Locked = True
    txtTitle.Locked = True
    txtDescription.Locked = True
    txtWebAddress.Locked = True
    txtShot.Locked = True
    cmdGetShot.Enabled = False

End Sub

Private Sub Call_DoMoveFiles(ByVal lNewCategoryNumber As Long, ByVal sOldCategoryPath As String, _
                             ByVal sNewCateroryPath As String, ByVal sTitleName As String)

  Dim rsFileDetails As Recordset
  Dim sFileName As String

    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    With rsFileDetails
        .FindFirst "TITLE = '" & sTitleName & "'"
        If Not .NoMatch Then
            sFileName = .Fields("FILE_NAME")
            Name sOldCategoryPath & "\" & sFileName As sNewCateroryPath & "\" & sFileName
            .Edit
            .Fields("CATEGORY_NUMBER") = lNewCategoryNumber
            .Update
            .Close
        End If
    End With 'RSFILEDETAILS
    Set rsFileDetails = Nothing

End Sub

Private Sub Call_DoRecView(ByVal sTitleToDisplay As String)

  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset

    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    With rsFileDetails
        .FindFirst "TITLE = '" & sTitleToDisplay & "'"
        If Not .NoMatch Then
            txtTitle.Text = Trim$(.Fields("TITLE"))
            txtDescription = Trim$(.Fields("DESCRIPTION"))
            If Not IsNull(.Fields("PAGE_ADDRESS")) Then
                txtWebAddress = Trim$(.Fields("PAGE_ADDRESS"))
            End If
        End If
        txtShot.Text = Trim$(.Fields("SCREENSHOT"))
        rsCategories.FindFirst "CATEGORY_NUMBER = " & .Fields("CATEGORY_NUMBER")
        If Not rsCategories.NoMatch Then
            txtCategory = rsCategories.Fields("CATEGORY_NAME")
        End If
        rsCategories.Close
        .Close
    End With 'RSFILEDETAILS
    Set rsFileDetails = Nothing
    Set rsCategories = Nothing

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

Private Sub Call_DoUnLockBoxes()

    txtTitle.Locked = False
    txtDescription.Locked = False
    txtWebAddress.Locked = False

End Sub

Private Sub Call_ThisFormSize()

    With Me
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

End Sub

Private Sub cmdCancel_Click()

    lstTitles.Enabled = True
    psOldTitle = vbNullString
    Call_DoBackCoulor &HFFFFFF 'White
    Call_DoLockBoxes
    Call_DoBoxClear
    optOneAll(0).Enabled = False
    optOneAll(1).Enabled = False
    cmdLoad.Enabled = True
    cmdMoveFiles.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    cmdGetShot.Enabled = False
    psButtonSelect = "Load"
    lstCategoryList_Click
    pbEdit = False

End Sub

Private Sub cmdCategorize_Click()

  Dim FSys As New FileSystemObject
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim rsCodeDay As Recordset
  Dim lCategoryNumber As Long
  Dim lCountRec As Long
  Dim sFoundTitleCategory As String
  Dim lFoundCategoryNumber As Long
  Dim sCategoryPath As String
  Dim lCountFoundTitles As Long

    Call_DoLockAll
    cmdLoad.Enabled = False
    cmdCategorize.Enabled = False
    lblTitlesFound.Caption = "0 Titles Found"
    Call_DoBackCoulor &HFFFFFF 'White
    lstTitles.BackColor = &HFFFFFF 'White
    lstTitles.Enabled = True
    pbAllCategories = True
    Screen.MousePointer = vbHourglass
    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    If rsCodeDay.RecordCount > 0 Then
        rsCodeDay.Close
        Set rsCodeDay = Nothing
        Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
        'We need to find the Category Number of the default Category.
        rsCategories.FindFirst "CATEGORY_NAME = '### Imported Files To Be Moved'"
        If Not rsCategories.NoMatch Then
            lCategoryNumber = rsCategories.Fields("CATEGORY_NUMBER")
            'Now we want only records with this Category Number from FILEDETAILS table
            Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails WHERE CATEGORY_NUMBER =" & lCategoryNumber)
            If rsFileDetails.RecordCount > 0 Then
                With rsFileDetails
                    .MoveFirst
                    .MoveLast
                    lCountRec = .RecordCount
                    .MoveFirst
                    For lCountRec = 1 To lCountRec
                        'Find out if the Title exists on CODEDAY table
                        sFoundTitleCategory = Func_CodeDayFindCategory(.Fields("TITLE"))
                        'If returns nothing (vbNullString) then is no Title on CODEDAY table
                        'Else going to look if for the Category Name
                        If sFoundTitleCategory <> vbNullString Then
                            rsCategories.FindFirst "CATEGORY_NAME = '" & sFoundTitleCategory & "'"
                            If Not rsCategories.NoMatch Then
                                'We have this Category on table
                                lFoundCategoryNumber = rsCategories.Fields("CATEGORY_NUMBER")
                                sCategoryPath = rsCategories.Fields("CATEGORY_PATH")
                                'Change the Category Number to be found the Category Name on Category table
                                .Edit
                                .Fields("CATEGORY_NUMBER") = lFoundCategoryNumber
                                .Update
                              Else 'NOT NOT...
                                'Category doesn't exist so must create a new Category on table
                                sCategoryPath = App.Path & "\ZipFiles\" & Func_FilterString(sFoundTitleCategory)
                                With rsCategories
                                    .AddNew
                                    .Fields("CATEGORY_NAME") = sFoundTitleCategory
                                    .Fields("CATEGORY_PATH") = sCategoryPath
                                    .Fields("CATEGORY_NUMBER") = .Fields("AutoNumber")
                                    lFoundCategoryNumber = .Fields("AutoNumber")
                                    .Update
                                    .MoveLast
                                End With 'rsCategories
                                'Change the Category Number to be found the Category Name on Category table
                                .Edit
                                .Fields("CATEGORY_NUMBER") = lFoundCategoryNumber
                                .Update
                                If Not FSys.FolderExists(sCategoryPath) Then
                                    FSys.CreateFolder sCategoryPath
                                End If
                            End If
                            'Move the file to a proper directory
                            'Tip from Jim Jose
                            Name App.Path & "\ZipFiles\### Imported Files To Be Moved\" _
                                 & .Fields("FILE_NAME") As sCategoryPath & "\" & .Fields("FILE_NAME")
                            'Delete record on CODEDAY table with this Title.
                            'We don't need it any more, suppose to be unique
                            Call_DoCodeDayDelTitle .Fields("TITLE")
                            Call_DoBoxClear
                            txtCategory.Text = sFoundTitleCategory
                            txtTitle.Text = .Fields("TITLE")
                            txtWebAddress.Text = .Fields("PAGE_ADDRESS")
                            txtShot.Text = .Fields("SCREENSHOT")
                            txtDescription.Text = .Fields("DESCRIPTION")
                            lstTitles.AddItem .Fields("TITLE")
                            lCountFoundTitles = lCountFoundTitles + 1
                            lblTitlesFound.Caption = lCountFoundTitles & " Titles Found"
                            DoEvents
                        End If
                        .MoveNext
                    Next lCountRec
                    rsCategories.Close
                    .Close
                End With 'RSFILEDETAILS
              Else 'NOT RSFILEDETAILS.RECORDCOUNT...
                rsFileDetails.Close
                MsgBox "No Tiles found on table." _
                       & vbNewLine & "Please import some ZIP files to be processed!" _
                       & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
            End If
            Set rsFileDetails = Nothing
          Else 'NOT NOT...
            rsCategories.Close
            MsgBox "The default Category doesn't exist!" _
                   & vbNewLine & "Please import some ZIP files to be processed!" _
                   & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
        End If
        Set rsCategories = Nothing
      Else 'NOT RSCODEDAY.RECORDCOUNT...
        rsCodeDay.Close
        Set rsCodeDay = Nothing
        MsgBox "Code of the Day table is empty!" _
               & vbNewLine & "Please import Code of the Day emails." _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    Screen.MousePointer = vbDefault
    cmdLoad.Enabled = True
    cmdCategorize.Enabled = True

End Sub

Private Sub cmdDelete_Click()

  Dim FSys As New FileSystemObject
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset

    pbEdit = True
    lstTitles.Enabled = False
    optOneAll(0).Enabled = False
    optOneAll(1).Enabled = False
    cmdLoad.Enabled = False
    cmdMoveFiles.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    txtTitle.BackColor = &HC0C0FF 'Light Red
    txtDescription.BackColor = &HC0C0FF 'Light Red
    txtWebAddress.BackColor = &HC0C0FF 'Light Red
    txtShot.BackColor = &HC0C0FF 'Light Red
    If Func_DoMBPositiveDel = vbYes Then
        With rsFileDetails
            Screen.MousePointer = vbHourglass
            .FindFirst "TITLE = '" & txtTitle.Text & "'"
            If Not .NoMatch Then
                rsCategories.FindFirst "CATEGORY_NUMBER = " & .Fields("CATEGORY_NUMBER")
                If Not rsCategories.NoMatch Then
                    If FSys.FileExists(rsCategories.Fields("CATEGORY_PATH") & "\" & .Fields("FILE_NAME")) Then
                        FSys.DeleteFile rsCategories.Fields("CATEGORY_PATH") & "\" & .Fields("FILE_NAME")
                    End If
                End If
                rsCategories.Close
                .Delete
                .Close
                Screen.MousePointer = vbDefault
            End If
        End With 'RSFILEDETAILS
        Call_DoMBBeenDel
    End If
    Set rsFileDetails = Nothing
    Set rsCategories = Nothing
    cmdCancel_Click

End Sub

Private Sub cmdEdit_Click()

    pbEdit = True
    lstTitles.Enabled = False
    optOneAll(0).Enabled = False
    optOneAll(1).Enabled = False
    psButtonSelect = vbNullString
    psOldTitle = txtTitle.Text
    psOldAddress = txtWebAddress.Text
    psOldShot = txtShot.Text
    cmdLoad.Enabled = False
    cmdMoveFiles.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    txtTitle.BackColor = &HC0FFC0 ' Light  Green
    txtDescription.BackColor = &HC0FFC0 ' Light  Green
    txtWebAddress.BackColor = &HC0FFC0 ' Light  Green
    Call_DoUnLockBoxes
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdGetShot.Enabled = True
    cmdSave.SetFocus

End Sub

Private Sub cmdGetShot_Click()

  Dim lCountChar As Long

    On Error Resume Next
        'Prompt for file opening
        With cdGetShot
            .DialogTitle = "Get New Screenshot"
            .Filter = "Picture Files(*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
            .ShowOpen
            psNewShotFileFromPath = .FileName
            .FileName = vbNullString
        End With 'CDGETSHOT
        If Len(psNewShotFileFromPath) > 0 Then
            If Left$(psNewShotFileFromPath, Len(App.Path)) = App.Path Then
                MsgBox "Files from the application path CANNOT be used." _
                       & vbNewLine & psNewShotFileFromPath _
                       & vbNewLine & "Please select other file." _
                       & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbCritical, gsLocalForm
                psNewShotFileFromPath = vbNullString
                txtShot.Text = vbNullString
              Else 'NOT LEFT$(PSNEWSHOTFILEFROMPATH,...
                lCountChar = Len(psNewShotFileFromPath)
                For lCountChar = lCountChar To 1 Step -1
                    If Mid$(psNewShotFileFromPath, lCountChar, 1) = "\" Then
                        txtShot.Text = Right$(psNewShotFileFromPath, (Len(psNewShotFileFromPath) - lCountChar))
                        Exit For 'loop varying lcountchar
                    End If
                Next lCountChar
            End If
          Else 'NOT LEN(PSNEWSHOTFILEFROMPATH)...
            MsgBox "The operation was cancelled!" _
                   & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
        End If
    On Error GoTo 0

End Sub

Private Sub cmdLoad_Click()

    Call_DoBoxClear
    lstTitles.Clear
    lstCategoryList.Clear
    Call_DoLockAll
    frameCategorylist.Top = frameRecResults.Top + 902
    psButtonSelect = "Load"
    psCategoryName = vbNullString
    frameCategorylist.Caption = "Category List - Load Files."
    Call_DoListCategory

End Sub

Private Sub cmdMainMenu_Click()

    Call_ThisFormSize
    frmStartMenu.Show
    Unload Me

End Sub

Private Sub cmdMoveFiles_Click()

    lblTitlesFound.Visible = False
    optOneAll(0).Enabled = False
    optOneAll(1).Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    frameCategorylist.Top = frameRecResults.Top + 902
    psButtonSelect = "MoveFiles"
    psCategoryName = txtCategory.Text
    If optOneAll(0).Value Then
        frameCategorylist.Caption = "Category List - Move One File."
      Else 'OPTONEALL(0).VALUE = FALSE/0
        frameCategorylist.Caption = "Category List - Move All Files."
    End If
    Call_DoListCategory

End Sub

Private Sub cmdSave_Click()

  Dim FSys As New FileSystemObject
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim sSearch As String
  Dim sNewFileName As String

    If Func_DoMBNewEditRec = vbNo Then
        cmdCancel_Click
      Else 'NOT FUNC_DOMBNEWEDITREC...
        Call_DoClearText
        If Func_DoBoxBlankCheck Then
            Exit Sub '---> Bottom
        End If
        If Func_LengthCount Then
            Exit Sub '---> Bottom
        End If
        Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
        Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
        'Find out if the new File Title already exists
        sSearch = Trim$(txtTitle.Text)
        If UCase$(psOldTitle) <> UCase$(sSearch) Then
            rsFileDetails.FindFirst "TITLE = '" & sSearch & "'"
            If Not rsFileDetails.NoMatch Then
                'The File Title has a duplicate so get out and change it
                Call_MBAlreadyExists sSearch, "File Title"
                txtTitle.SetFocus
                Exit Sub '---> Bottom
            End If
        End If
        'Find out if the new Web Address already exists
        sSearch = Trim$(txtWebAddress.Text)
        If UCase$(psOldAddress) <> UCase$(sSearch) Then
            rsFileDetails.FindFirst "PAGE_ADDRESS = '" & sSearch & "'"
            If Not rsFileDetails.NoMatch Then
                'The Web Address has a duplicate so get out and change it
                Call_MBAlreadyExists sSearch, "Web Address"
                txtWebAddress.SetFocus
                Exit Sub '---> Bottom
            End If
        End If
        sSearch = Trim$(txtShot.Text)
        If UCase$(psOldShot) <> UCase$(sSearch) Then
            rsFileDetails.FindFirst "SCREENSHOT = '" & sSearch & "'"
            If Not rsFileDetails.NoMatch Then
                'The Screenshot has a duplicate so get out and change it
                Call_MBAlreadyExists sSearch, "Screenshot"
                Exit Sub '---> Bottom
            End If
        End If
        'Is no duplicates let's find where it is the record with the old Title and change the details
        sNewFileName = Func_FilterString(Trim$(txtTitle.Text))
        sNewFileName = sNewFileName & ".zip"
        With rsFileDetails
            .FindFirst "TITLE = '" & psOldTitle & "'"
            If Not .NoMatch Then
                rsCategories.FindFirst "CATEGORY_NUMBER = " & .Fields("CATEGORY_NUMBER")
                If Not rsCategories.NoMatch Then
                    If FSys.FileExists(rsCategories.Fields("CATEGORY_PATH") & "\" & .Fields("FILE_NAME")) Then
                        FSys.MoveFile rsCategories.Fields("CATEGORY_PATH") & "\" & .Fields("FILE_NAME"), _
                                                          rsCategories.Fields("CATEGORY_PATH") & "\" & sNewFileName
                    End If
                End If
                rsCategories.Close
                If UCase$(psOldShot) <> UCase$(Trim$(txtShot.Text)) Then
                    If Len(psNewShotFileFromPath) > 0 Then
                        If FSys.FileExists(App.Path & "\ScreenshotPics\" & psOldShot) Then
                            FSys.DeleteFile App.Path & "\ScreenshotPics\" & psOldShot
                        End If
                        If Not FSys.FileExists(App.Path & "\ScreenshotPics\" & txtShot.Text) Then
                            FSys.CopyFile psNewShotFileFromPath, App.Path & "\ScreenshotPics\" & txtShot.Text
                        End If
                    End If
                End If
                .Edit
                .Fields("TITLE") = Trim$(txtTitle.Text)
                .Fields("DESCRIPTION") = Trim$(txtDescription.Text)
                .Fields("FILE_NAME") = sNewFileName
                .Fields("PAGE_ADDRESS") = Trim$(txtWebAddress.Text) & vbNullString
                .Fields("SCREENSHOT") = Trim$(txtShot.Text) & vbNullString
                .Update
                .Close
            End If
        End With 'RSFILEDETAILS
        Call_DoMBEditUpdate
        cmdCancel_Click
        Set rsFileDetails = Nothing
    End If

End Sub

Private Sub Form_Load()

  Dim rsCategories As Recordset

    gsLocalForm = Me.Caption
    Me.Caption = gsProgName & " - " & Me.Caption & " - " & gsOwner
    With Me
        .Height = glFormHeight
        .Left = glFormLeft
        .Top = glFormTop
        .Width = glFormWidth
    End With 'ME
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    If rsCategories.RecordCount < 1 Then
        cmdLoad.Enabled = False
    End If
    Call_DoLockAll
    Call_DoToolTips

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

Private Function Func_CodeDayFindCategory(ByVal sFindTitle As String) As String

  Dim rsCodeDay As Recordset

    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    With rsCodeDay
        .FindFirst "CODE_TITLE = '" & sFindTitle & "'"
        If Not .NoMatch Then
            Func_CodeDayFindCategory = .Fields("CODE_CATEGORY")
          Else 'NOT NOT...
            Func_CodeDayFindCategory = vbNullString
        End If
        .Close
    End With 'RSCODEDAY
    Set rsCodeDay = Nothing

End Function

Private Function Func_DoBoxBlankCheck() As Boolean

  'Find if the string is Null or contains only spaces

    Select Case vbNullString
      Case Trim$(txtTitle.Text)
        If Func_BoxBlank("File Title") Then
            Func_DoBoxBlankCheck = True
        End If
        txtTitle.SetFocus
      Case Trim$(txtDescription.Text)
        If Func_BoxBlank("File Description") Then
            Func_DoBoxBlankCheck = True
        End If
        txtDescription.SetFocus
    End Select

End Function

Private Function Func_LengthCount() As Boolean

  'Counting characters to find if have more the maximum
  'I set the records to the maximum (255) if you would like to change it
  'do not forget to change here too.
  'The Memo field is limited to 32000 because, I think,
  'Windows 9x dues not support more than that

    Select Case True
      Case Func_MaxBoxLength(txtTitle.Text, "File Title", 255)
        Func_LengthCount = True
        txtTitle.SetFocus
      Case Func_MaxBoxLength(txtDescription.Text, "File Description", 32000)
        Func_LengthCount = True
        txtDescription.SetFocus
      Case Func_MaxBoxLength(txtWebAddress.Text, "Web Address", 255)
        Func_LengthCount = True
        txtWebAddress.SetFocus
      Case Func_MaxBoxLength(txtShot.Text, "Screenshot", 255)
        Func_LengthCount = True
        'txtShot.SetFocus
    End Select

End Function

Private Sub lstCategoryList_Click()

  Dim rsCategories As Recordset
  Dim rsFileDetails As Recordset
  Dim lOldCategoryNumber As Long
  Dim lNewCategoryNumber As Long
  Dim sOldCategoryPath As String
  Dim sNewCateroryPath As String
  Dim lCountList As Long
  Dim sListTitleName As String
  Dim sSelectedCategory As String
  Dim lCountMovedFiles As Long
  Dim sNewCateroryName As String

    Screen.MousePointer = vbHourglass
    Select Case True
      Case psButtonSelect = "Load"
        If pbEdit Then
            sSelectedCategory = psCurrentCategory
          Else 'PBEDIT = FALSE/0
            psCurrentCategory = lstCategoryList.Text
            sSelectedCategory = psCurrentCategory
        End If
        Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
        With rsCategories
            .FindFirst "CATEGORY_NAME = '" & sSelectedCategory & "'"
            If Not .NoMatch Then
                lOldCategoryNumber = .Fields("CATEGORY_NUMBER")
            End If
            .Close
        End With 'RSCATEGORIES
        Call_DoListTitles lOldCategoryNumber, sSelectedCategory
        frameCategorylist.Top = 22000
      Case psButtonSelect = "MoveFiles"
        Select Case True
          Case optOneAll(0).Value
            frameCategorylist.Top = 22000
            DoEvents
            Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
            With rsCategories
                .FindFirst "CATEGORY_NAME = '" & psCategoryName & "'"
                If Not .NoMatch Then
                    lOldCategoryNumber = .Fields("CATEGORY_NUMBER")
                    sOldCategoryPath = .Fields("CATEGORY_PATH")
                End If
                .FindFirst "CATEGORY_NAME = '" & lstCategoryList.Text & "'"
                If Not .NoMatch Then
                    lNewCategoryNumber = .Fields("CATEGORY_NUMBER")
                    sNewCateroryPath = .Fields("CATEGORY_PATH")
                End If
                .Close
            End With 'RSCATEGORIES
            Call_DoMoveFiles lNewCategoryNumber, sOldCategoryPath, sNewCateroryPath, txtTitle.Text
            lCountList = lstTitles.ListCount - 1
            For lCountList = 0 To lCountList
                If lstTitles.List(lCountList) = txtTitle.Text Then
                    lstTitles.RemoveItem (lCountList)
                End If
            Next lCountList
            lstTitles.Refresh
            Call_DoBoxClear
          Case optOneAll(1).Value
            frameCategorylist.Top = 22000
            txtTitle.Text = vbNullString
            txtCategory.Text = vbNullString
            txtWebAddress.Text = vbNullString
            txtDescription.Text = vbNullString
            txtShot.Text = vbNullString
            DoEvents
            lCountList = lstTitles.ListCount - 1
            For lCountList = 0 To lCountList
                sListTitleName = lstTitles.List(0)
                Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails")
                Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
                With rsFileDetails
                    .FindFirst "TITLE = '" & sListTitleName & "'"
                    If Not .NoMatch Then
                        lOldCategoryNumber = .Fields("CATEGORY_NUMBER")
                        .Close
                    End If
                End With 'RSFILEDETAILS
                With rsCategories
                    .FindFirst "CATEGORY_NUMBER = " & lOldCategoryNumber
                    If Not .NoMatch Then
                        sOldCategoryPath = .Fields("CATEGORY_PATH")
                    End If
                    .FindFirst "CATEGORY_NAME = '" & lstCategoryList.Text & "'"
                    If Not .NoMatch Then
                        lNewCategoryNumber = .Fields("CATEGORY_NUMBER")
                        sNewCateroryPath = .Fields("CATEGORY_PATH")
                        sNewCateroryName = .Fields("CATEGORY_NAME")
                    End If
                    .Close
                End With 'RSCATEGORIES
                Call_DoMoveFiles lNewCategoryNumber, sOldCategoryPath, sNewCateroryPath, sListTitleName
                lCountMovedFiles = lCountMovedFiles + 1
                txtTitle.Text = sListTitleName
                txtCategory.Text = lCountMovedFiles & "  Files Moved to " & sNewCateroryName & " Category"
                DoEvents
                lstTitles.RemoveItem (0)
                lstTitles.Refresh
            Next lCountList
        End Select
    End Select
    lstCategoryList.Clear
    Set rsCategories = Nothing
    Set rsFileDetails = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub lstCategoryList_LostFocus()

    lstCategoryList.Clear
    frameCategorylist.Top = 22000

End Sub

Private Sub lstTitles_Click()

    If pbAllCategories Then
        optOneAll(0).Value = True
        optOneAll(0).Enabled = True
        optOneAll(1).Enabled = False
      Else 'PBALLCATEGORIES = FALSE/0
        optOneAll(0).Value = True
        optOneAll(0).Enabled = True
        optOneAll(1).Enabled = True
    End If
    cmdMoveFiles.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    Call_DoBoxClear
    Call_DoRecView lstTitles.Text

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

':)Code Fixer V3.0.9 (04/08/2005 18:02:32) 11 + 939 = 950 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 11  Code: 968  Total: 979 Lines
':) CommentOnly: 31 (3.2%)  Commented: 58 (5.9%)  Empty: 136 (13.9%)  Max Logic Depth: 9
