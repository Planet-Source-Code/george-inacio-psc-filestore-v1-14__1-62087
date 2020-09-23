VERSION 5.00
Begin VB.Form frmCategories 
   Caption         =   "Add, Edit and Delete Categories"
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
      TabIndex        =   6
      ToolTipText     =   "Returns to Main Menu."
      Top             =   6225
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
      TabIndex        =   1
      ToolTipText     =   "Stop the current operation."
      Top             =   3975
      Width           =   1500
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      ToolTipText     =   "Click to Add new record the Category table"
      Top             =   4350
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
      ToolTipText     =   $"frmCategories.frx":0000
      Top             =   4725
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
      ToolTipText     =   "Save the record to the Database."
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
      TabIndex        =   5
      ToolTipText     =   "Deletes the selected entry.|Note: Directory and all files in the directory|will be erased and CANNOT be recover."
      Top             =   5475
      Width           =   1500
   End
   Begin VB.Frame frameCategories 
      Caption         =   "Category Names"
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
      Height          =   2850
      Left            =   225
      TabIndex        =   8
      Top             =   3720
      Width           =   9000
      Begin VB.ListBox lstCategories 
         Height          =   2400
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Click an entry to view the record."
         Top             =   270
         Width           =   8750
      End
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
      TabIndex        =   7
      Top             =   100
      Width           =   10695
      Begin VB.PictureBox picCFXPBugFixfrmEditMove 
         BorderStyle     =   0  'None
         Height          =   3038
         Left            =   100
         ScaleHeight     =   3045
         ScaleWidth      =   10500
         TabIndex        =   9
         Top             =   276
         Width           =   10495
         Begin VB.TextBox txtCategoryPath 
            Height          =   285
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   $"frmCategories.frx":00CB
            Top             =   2200
            Width           =   8500
         End
         Begin VB.TextBox txtCategoryName 
            Height          =   285
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Type the name of the Category here.|If directory dues not exist will be created."
            Top             =   1300
            Width           =   8500
         End
         Begin VB.TextBox txtCategoryNumber 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "The number that has been automatically assigned to this category.|Note: This is assigned automatically and CANOT be changed."
            Top             =   400
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Category Path"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   140
            TabIndex        =   12
            Top             =   2200
            Width           =   1700
         End
         Begin VB.Label Label2 
            Caption         =   "Category Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   140
            TabIndex        =   11
            Top             =   1300
            Width           =   1700
         End
         Begin VB.Label Label1 
            Caption         =   "Category Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   285
            Left            =   140
            TabIndex        =   10
            Top             =   400
            Width           =   1700
         End
      End
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
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tooltips As New Collection
Private pbAdd As Boolean
Private pbEdit As Boolean
Private sOldCategoryName As String
Private sOldCategoryPath As String

Private Sub Call_DoBoxClear()

    txtCategoryNumber.Text = vbNullString
    txtCategoryName.Text = vbNullString
    txtCategoryPath.Text = vbNullString

End Sub

Private Sub Call_DoBoxLock()

    txtCategoryName.Locked = True

End Sub

Private Sub Call_DoBoxUnLock()

    txtCategoryName.Locked = False

End Sub

Private Sub Call_DoCategoryList()

  Dim rsCategories As Recordset
  Dim lCountRecords As Long

    lstCategories.Clear
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    If rsCategories.RecordCount > 0 Then
        cmdCancel.Enabled = False
        cmdEdit.Enabled = False
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        cmdAdd.Enabled = True
        txtCategoryNumber.BackColor = &HFFFFFF 'White
        txtCategoryPath.BackColor = &HFFFFFF 'White
        txtCategoryName.BackColor = &HFFFFFF 'White
        With rsCategories
            .MoveFirst
            .MoveLast
            lCountRecords = .RecordCount
            .MoveFirst
            For lCountRecords = 1 To lCountRecords
                lstCategories.AddItem .Fields("CATEGORY_NAME")
                .MoveNext
            Next lCountRecords
            .Close
        End With 'RSCATEGORIES
      Else 'NOT RSCATEGORIES.RECORDCOUNT...
        Call_DoEmptyDBCheck
    End If
    Set rsCategories = Nothing

End Sub

Private Sub Call_DoClearText()

  'Call Function to replace Single Quote with one Apostrophe and Double Quotes with two Apostrophes

    txtCategoryName.Text = Func_SrchReplace(txtCategoryName.Text)

End Sub

Private Sub Call_DoEmptyDBCheck()

    With lstCategories
        .Clear
        .BackColor = &H8000000F  'Grey
        .AddItem vbNullString
        .AddItem "No Entries Found"
        .AddItem "Table is Empty! Please Click The ADD Button To Begin Entering Data."
    End With 'LSTCATEGORIES
    txtCategoryNumber.BackColor = &H8000000F    'Grey
    txtCategoryPath.BackColor = &H8000000F    'Grey
    txtCategoryName.BackColor = &H8000000F    'Grey
    cmdCancel.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdAdd.Enabled = True
    Call_DoMBDatabaseEmpty

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
    Set frmCategories = Nothing

End Sub

Private Sub Call_DoSaveAdd()

  Dim FSys As New FileSystemObject
  Dim rsCategories As Recordset

    If Func_DoMBAddNewRec = vbYes Then
        Call_DoClearText
        If Func_DoBoxBlankCheck Then
            Exit Sub '---> Bottom
        End If
        If Func_LengthCount Then
            Exit Sub '---> Bottom
        End If
        Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
        rsCategories.FindFirst "CATEGORY_NAME = '" & Trim$(txtCategoryName.Text) & "'"
        If Not rsCategories.NoMatch Then
            Call_MBAlreadyExists txtCategoryName.Text, "Category Name"
            rsCategories.Close
            Set rsCategories = Nothing
            Exit Sub '---> Bottom
        End If
        If FSys.FolderExists(Trim$(txtCategoryPath.Text)) Then
            Call_MBAlreadyExists txtCategoryPath.Text, "Category Path"
            Exit Sub '---> Bottom
          Else 'FSYS.FOLDEREXISTS(TRIM$(TXTCATEGORYPATH.TEXT)) = FALSE/0
            FSys.CreateFolder (Trim$(txtCategoryPath.Text))
        End If
        With rsCategories
            .AddNew
            .Fields("CATEGORY_NUMBER") = .Fields("AutoNumber")
            .Fields("CATEGORY_NAME") = Trim$(txtCategoryName.Text)
            .Fields("CATEGORY_PATH") = Trim$(txtCategoryPath.Text)
            .Update
            .MoveLast
            .Close
        End With 'RSCATEGORIES
        Set rsCategories = Nothing
        Call_DoMBNewRecAdded
        cmdCancel_Click
    End If

End Sub

Private Sub Call_DoSaveEdit()

  Dim FSys As New FileSystemObject
  Dim rsCategories As Recordset
  Dim sSearch As String

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
        Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
        sSearch = Trim$(txtCategoryName.Text)
        If UCase$(sSearch) <> UCase$(sOldCategoryName) Then
            rsCategories.FindFirst "CATEGORY_NAME = '" & Trim$(txtCategoryName.Text) & "'"
            If Not rsCategories.NoMatch Then
                Call_MBAlreadyExists txtCategoryName.Text, "Category Name"
                rsCategories.Close
                Set rsCategories = Nothing
                Exit Sub '---> Bottom
            End If
        End If
        sSearch = Trim$(txtCategoryPath.Text)
        If UCase$(sSearch) <> UCase$(sOldCategoryPath) Then
            If FSys.FolderExists(Trim$(txtCategoryPath.Text)) Then
                Call_MBAlreadyExists txtCategoryPath.Text, "Category Path"
                Exit Sub '---> Bottom
            End If
        End If
        If Not FSys.FolderExists(Trim$(txtCategoryPath.Text)) Then
            FSys.MoveFolder sOldCategoryPath, Trim$(txtCategoryPath.Text)
        End If
        With rsCategories
            .FindFirst "CATEGORY_NAME = '" & sOldCategoryName & "'"
            If Not .NoMatch Then
                .Edit
                .Fields("CATEGORY_NAME") = Trim$(txtCategoryName.Text)
                .Fields("CATEGORY_PATH") = Trim$(txtCategoryPath.Text)
                .Update
                .Close
            End If
        End With 'RSCATEGORIES
        Set rsCategories = Nothing
        Call_DoMBNewRecAdded
        cmdCancel_Click
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

Private Sub cmdAdd_Click()

    pbAdd = True
    lstCategories.Clear
    Call_DoBoxClear
    Call_DoBoxUnLock
    txtCategoryName.BackColor = &HC0FFFF ' Light  Yellow
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    txtCategoryNumber = vbNullString
    txtCategoryPath = App.Path & "\ZipFiles\"
    txtCategoryName.SetFocus

End Sub

Private Sub cmdCancel_Click()

    Call_DoBoxClear
    Call_DoBoxUnLock
    txtCategoryName.BackColor = &HFFFFFF 'White
    cmdCancel.Enabled = False
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    pbAdd = False
    pbEdit = False
    lstCategories.Enabled = True
    Call_DoCategoryList

End Sub

Private Sub cmdDelete_Click()

  Dim FSys As New FileSystemObject
  Dim rsFileDetails As Recordset
  Dim rsCategories As Recordset
  Dim lCountRecords As Long

    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    txtCategoryName.BackColor = &HC0C0FF 'Light Red
    txtCategoryNumber.BackColor = &HC0C0FF 'Light Red
    txtCategoryPath.BackColor = &HC0C0FF 'Light Red
    Set rsFileDetails = DB1.OpenRecordset("SELECT * FROM FileDetails WHERE CATEGORY_NUMBER = " & txtCategoryNumber.Text)
    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    If Func_DoMBPositiveDel = vbYes Then
        If Func_DoMBSureDel("Directory and All files") = vbYes Then
            Screen.MousePointer = vbHourglass
            If FSys.FolderExists(txtCategoryPath.Text) Then
                FSys.DeleteFolder txtCategoryPath.Text
            End If
            If rsFileDetails.RecordCount > 0 Then
                With rsFileDetails
                    .MoveFirst
                    .MoveLast
                    lCountRecords = .RecordCount
                    .MoveFirst
                    For lCountRecords = 1 To lCountRecords
                        .Delete
                        .MoveNext
                    Next lCountRecords
                    .Close
                End With 'RSFILEDETAILS
            End If
            If rsCategories.RecordCount > 0 Then
                With rsCategories
                    .FindFirst "CATEGORY_NUMBER = " & txtCategoryNumber.Text
                    If Not .NoMatch Then
                        .Delete
                        .Close
                    End If
                End With 'RSCATEGORIES
            End If
            Screen.MousePointer = vbDefault
            Call_DoMBBeenDel
        End If
    End If
    Set rsFileDetails = Nothing
    Set rsCategories = Nothing
    cmdCancel_Click

End Sub

Private Sub cmdEdit_Click()

    pbEdit = True
    lstCategories.Enabled = False
    sOldCategoryName = Trim$(txtCategoryName.Text)
    sOldCategoryPath = Trim$(txtCategoryPath.Text)
    Call_DoBoxUnLock
    txtCategoryName.BackColor = &HC0FFC0 ' Light  Green
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    txtCategoryName.SetFocus

End Sub

Private Sub cmdMainMenu_Click()

    Call_ThisFormSize
    frmStartMenu.Show
    Unload Me

End Sub

Private Sub cmdSave_Click()

    Select Case True
      Case pbAdd
        Call_DoSaveAdd
        pbAdd = False
      Case pbEdit
        Call_DoSaveEdit
        pbEdit = False
    End Select

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
    Call_DoToolTips
    Call_DoCategoryList

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

Private Function Func_DoBoxBlankCheck() As Boolean

  'Find if the string is Null or contains only spaces

    Select Case vbNullString
      Case Trim$(txtCategoryName.Text)
        If Func_BoxBlank("Caterory Name") Then
            Func_DoBoxBlankCheck = True
        End If
        txtCategoryName.SetFocus
    End Select

End Function

Private Function Func_LengthCount() As Boolean

  'Counting characters to find if have more the maximum
  'I set the records to the maximum (255) if you would like to change it
  'do not forget to change here too.

    Select Case True
      Case Func_MaxBoxLength(txtCategoryName.Text, "Caterory Name", 255)
        Func_LengthCount = True
        txtCategoryName.SetFocus
      Case Func_MaxBoxLength(txtCategoryPath.Text, "Caterory Path", 255)
        Func_LengthCount = True
        txtCategoryPath.SetFocus
    End Select

End Function

Private Sub lstCategories_Click()

  Dim rsCategories As Recordset

    Set rsCategories = DB1.OpenRecordset("SELECT * FROM Categories")
    If rsCategories.RecordCount > 0 Then
        With rsCategories
            .FindFirst "CATEGORY_NAME = '" & lstCategories.Text & "' "
            If Not .NoMatch Then
                txtCategoryNumber.Text = .Fields("CATEGORY_NUMBER")
                txtCategoryName.Text = .Fields("CATEGORY_NAME")
                txtCategoryPath.Text = .Fields("CATEGORY_PATH")
                .Close
            End If
        End With 'RSCATEGORIES
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        Call_DoBoxLock
        lstCategories.SetFocus
    End If
    Set rsCategories = Nothing

End Sub

Private Sub mnuAboutItem_Click()

    frmAbout.Show vbModal

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

Private Sub txtCategoryName_Change()

  Dim sDirectoryName As String
  'Filtering the Category Path for directories illegal characters

    sDirectoryName = Func_SrchReplace(Trim$(txtCategoryName.Text))
    sDirectoryName = Func_FilterString(sDirectoryName)
    txtCategoryPath = App.Path & "\ZipFiles\" & sDirectoryName

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:34) 6 + 541 = 547 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 6  Code: 517  Total: 523 Lines
':) CommentOnly: 9 (1.7%)  Commented: 36 (6.9%)  Empty: 106 (20.3%)  Max Logic Depth: 6
