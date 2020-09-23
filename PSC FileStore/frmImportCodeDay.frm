VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmImportCodeDay 
   Caption         =   "Import Code of the Day"
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
   Begin VB.Frame frameDeleteDate 
      Caption         =   "Delete Titles by Date range"
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
      Height          =   2900
      Left            =   3000
      TabIndex        =   10
      Top             =   21900
      Width           =   4300
      Begin VB.PictureBox picCFXPBugFixfrmImportCodeDay 
         BorderStyle     =   0  'None
         Height          =   2563
         Index           =   7
         Left            =   100
         ScaleHeight     =   2565
         ScaleWidth      =   4095
         TabIndex        =   11
         Top             =   276
         Width           =   4100
         Begin MSACAL.Calendar Calendar1 
            Height          =   2535
            Left            =   20
            TabIndex        =   12
            Top             =   20000
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   4471
            _StockProps     =   1
            BackColor       =   -2147483633
            Year            =   2005
            Month           =   1
            Day             =   22
            DayLength       =   1
            MonthLength     =   2
            DayFontColor    =   0
            FirstDay        =   1
            GridCellEffect  =   1
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   0   'False
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdDeleteDate 
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
            Height          =   495
            Left            =   2195
            TabIndex        =   16
            ToolTipText     =   "Click to initiate the Delete records|Between the two Dates."
            Top             =   1902
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelDate 
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
            Height          =   495
            Left            =   620
            TabIndex        =   15
            ToolTipText     =   "Click to Cancel the operation."
            Top             =   1902
            Width           =   1335
         End
         Begin VB.TextBox txtFromDate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1700
            Locked          =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Click the box to select the Delete Starting date from the Calendar."
            Top             =   522
            Width           =   1700
         End
         Begin VB.TextBox txtToDate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1700
            Locked          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Click the box to select the Delete Ending date from the Calendar."
            Top             =   927
            Width           =   1700
         End
         Begin VB.Label lblToDate 
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   560
            TabIndex        =   18
            Top             =   927
            Width           =   1020
         End
         Begin VB.Label lblFromDate 
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   560
            TabIndex        =   17
            Top             =   522
            Width           =   1020
         End
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
      Left            =   8920
      TabIndex        =   5
      ToolTipText     =   "Returns to Main Menu."
      Top             =   5850
      Width           =   2000
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Emails"
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
      Left            =   8920
      TabIndex        =   1
      ToolTipText     =   $"frmImportCodeDay.frx":0000
      Top             =   3975
      Width           =   2000
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Titles"
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
      Left            =   8920
      TabIndex        =   2
      ToolTipText     =   "Displays all the Titles on the listbox Title Names."
      Top             =   4350
      Width           =   2000
   End
   Begin VB.CommandButton cmdDeleteTitle 
      Caption         =   "Delete Title"
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
      Left            =   8920
      TabIndex        =   3
      ToolTipText     =   "Select only one Title to be deleted."
      Top             =   4725
      Width           =   2000
   End
   Begin VB.CommandButton cmdDeleteByDate 
      Caption         =   "Delete by Date"
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
      Left            =   8920
      TabIndex        =   4
      ToolTipText     =   $"frmImportCodeDay.frx":0088
      Top             =   5100
      Width           =   2000
   End
   Begin VB.Frame frameTitles 
      Caption         =   "Title Names"
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
      TabIndex        =   7
      Top             =   3720
      Width           =   8500
      Begin VB.ListBox lstTitle 
         Height          =   2400
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Click to select the Title details.|Must click the Display button to show all Titles."
         Top             =   270
         Width           =   8250
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
      Height          =   2300
      Left            =   225
      TabIndex        =   6
      Top             =   300
      Width           =   10695
      Begin VB.PictureBox picCFXPBugFixfrmImportCodeDay 
         BorderStyle     =   0  'None
         Height          =   1963
         Index           =   0
         Left            =   100
         ScaleHeight     =   1965
         ScaleWidth      =   10500
         TabIndex        =   19
         Top             =   276
         Width           =   10495
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   2720
            Locked          =   -1  'True
            TabIndex        =   23
            ToolTipText     =   "Title found on the email."
            Top             =   957
            Width           =   7215
         End
         Begin VB.TextBox txtCategory 
            Height          =   285
            Left            =   2720
            Locked          =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "Category found on the email."
            Top             =   552
            Width           =   7215
         End
         Begin VB.TextBox txtWebAddress 
            Height          =   285
            Left            =   2720
            Locked          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Author Web Page at PSC found on the email."
            Top             =   1362
            Width           =   7215
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2710
            Locked          =   -1  'True
            TabIndex        =   20
            ToolTipText     =   "Code of the Day email date."
            Top             =   142
            Width           =   7215
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
            ForeColor       =   &H00800080&
            Height          =   285
            Left            =   1020
            TabIndex        =   27
            Top             =   1362
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Title Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Left            =   1020
            TabIndex        =   26
            Top             =   957
            Width           =   1695
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
            ForeColor       =   &H00800080&
            Height          =   285
            Left            =   1020
            TabIndex        =   25
            Top             =   552
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Email Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Left            =   1020
            TabIndex        =   24
            Top             =   142
            Width           =   1700
         End
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   "0 Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   400
      Left            =   2452
      TabIndex        =   9
      Top             =   2950
      Width           =   3000
   End
   Begin VB.Label lblEmails 
      Caption         =   "0 Emails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   5692
      TabIndex        =   8
      Top             =   2955
      Width           =   3000
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
Attribute VB_Name = "frmImportCodeDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tooltips As New Collection
Private plNumberOfTitles As Long
Private plNumberOfEmails As Long
Private plCalendar As Long

Private Sub Calendar1_Click()

    Select Case True
      Case plCalendar = 1
        txtFromDate.Text = Format$(Calendar1.Value, "dd mmmm yyyy")
        txtFromDate.SetFocus
      Case plCalendar = 2
        txtToDate.Text = Format$(Calendar1.Value, "dd mmmm yyyy")
        txtToDate.SetFocus
    End Select
    Calendar1.Top = 20240
    plCalendar = 0

End Sub

Private Sub Call_DoBoxClear()

    txtDate.Text = vbNullString
    txtCategory.Text = vbNullString
    txtTitle.Text = vbNullString
    txtWebAddress.Text = vbNullString
    txtFromDate.Text = vbNullString
    txtToDate.Text = vbNullString

End Sub

Private Sub Call_DoClearText()

  'Call Function to replace Single Quote with one Apostrophe and Double Quotes with two Apostrophes

    txtCategory.Text = Func_SrchReplace(txtCategory.Text)
    txtTitle.Text = Func_SrchReplace(txtTitle.Text)
    txtWebAddress.Text = Func_SrchReplace(txtWebAddress.Text)

End Sub

Private Sub Call_DoEmptyTable()

  Dim rsCodeDay As Recordset

    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    If rsCodeDay.RecordCount > 0 Then
        cmdDisplay.Enabled = True
      Else 'NOT RSCODEDAY.RECORDCOUNT...
        cmdDisplay.Enabled = False
        MsgBox "Code of the Day table is empty!" _
               & vbNewLine & "Import some Emails." _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    rsCodeDay.Close
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
    Set frmImportCodeDay = Nothing

End Sub

Private Sub Call_DoImportEmails(ByVal sDirectoryPath As String)

  Dim sFullPath() As String
  Dim lNumElements As Long
  Dim sFileName As String

    ReDim sFullPath(0) As String
    sFileName = Dir(sDirectoryPath & "*.eml", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    Do While Len(sFileName) <> 0
        lNumElements = UBound(sFullPath()) + 1
        ReDim Preserve sFullPath(lNumElements)
        sFullPath(lNumElements) = sDirectoryPath & sFileName
        sFileName = Dir()  ' Get next file.
    Loop
    For lNumElements = 1 To lNumElements
        Call_DoReadEml sFullPath(lNumElements)
        plNumberOfEmails = plNumberOfEmails + 1
        lblEmails.Caption = plNumberOfEmails & " Emails"
        DoEvents
    Next lNumElements

End Sub

Private Sub Call_DoReadEml(ByVal sEmlPath As String)

  Dim rsCodeDay As Recordset
  Dim FSys As New FileSystemObject
  Dim tsFile As TextStream
  Dim sLineText As String
  Dim bFoundDate As Boolean
  Dim bIsNumber As Boolean
  Dim bIsCategory As Boolean
  Dim bIsWebAddress As Boolean

    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    Set tsFile = FSys.OpenTextFile(sEmlPath, ForReading, False)
    With tsFile
        Do While .AtEndOfStream <> True
            sLineText = .ReadLine
            'Get the Email Date
            If Not bFoundDate Then
                If Func_WeekDays(sLineText) Then
                    txtDate.Text = Format$(Func_GetDate(sLineText), "dd mmmm yyyy")
                    bFoundDate = True
                End If
            End If
            'Get the Title, the one I am looking for have NO Space after the ).
            If Not bIsNumber Then
                If IsNumeric(Left$(sLineText, 1)) Then
                    If Func_IsSpace(sLineText) Then
                        txtTitle.Text = Trim$(Func_GetTitle(sLineText))
                        bIsNumber = True
                    End If
                End If
            End If
            'Get the Category name
            If Not bIsCategory Then
                If Left$(sLineText, 10) = "Category: " Then
                    txtCategory.Text = Trim$(Right$(sLineText, (Len(sLineText) - 10)))
                    bIsCategory = True
                End If
            End If
            'Get the Author Web Page at PSC
            If Not bIsWebAddress Then
                If Trim$(sLineText) = "Complete source code is at:" Then
                    sLineText = .ReadLine
                    txtWebAddress.Text = Trim$(sLineText)
                    bIsCategory = False
                    bIsNumber = False
                    Call_DoClearText
                    If Len(txtTitle.Text) > 0 Then
                        If Len(txtCategory.Text) > 0 Then
                            With rsCodeDay
                                If .RecordCount > 0 Then
                                    .FindFirst "CODE_TITLE = '" & txtTitle.Text & "'"
                                    If .NoMatch Then
                                        .AddNew
                                        .Fields("CODE_DATE") = txtDate.Text
                                        .Fields("CODE_CATEGORY") = txtCategory.Text
                                        .Fields("CODE_TITLE") = txtTitle.Text
                                        .Fields("CODE_WEB_ADDRESS") = txtWebAddress.Text
                                        .Update
                                        .MoveLast
                                        plNumberOfTitles = plNumberOfTitles + 1
                                        lblTitle.Caption = plNumberOfTitles & " Titles"
                                        DoEvents
                                    End If
                                  Else 'NOT .RECORDCOUNT...
                                    .AddNew
                                    .Fields("CODE_DATE") = txtDate.Text
                                    .Fields("CODE_CATEGORY") = txtCategory.Text
                                    .Fields("CODE_TITLE") = txtTitle.Text
                                    .Fields("CODE_WEB_ADDRESS") = txtWebAddress.Text
                                    .Update
                                    .MoveLast
                                    plNumberOfTitles = plNumberOfTitles + 1
                                    lblTitle.Caption = plNumberOfTitles & " Titles"
                                    DoEvents
                                End If
                            End With 'RSCODEDAY
                        End If
                    End If
                    DoEvents
                End If
            End If
        Loop
        .Close
        rsCodeDay.Close
    End With 'TSFILE
    Set rsCodeDay = Nothing

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

Private Sub cmdCancelDate_Click()

    frameDeleteDate.Top = 21900
    lstTitle.Clear
    cmdDeleteTitle.Enabled = False
    cmdDeleteByDate.Enabled = False
    Call_DoBoxClear
    Call_DoEmptyTable

End Sub

Private Sub cmdDeleteByDate_Click()

    frameDeleteDate.Top = 1900
    cmdDeleteDate.Enabled = False

End Sub

Private Sub cmdDeleteDate_Click()

  Dim rsCodeDay As Recordset
  Dim daFromDate As Date
  Dim daToDate As Date
  Dim lCountRec As Long

    daFromDate = txtFromDate.Text
    daToDate = txtToDate.Text
    If daFromDate > daToDate Then
        MsgBox "From Date is Grater then To Date!" _
               & vbNewLine & "Please Try A Different Date." _
               & vbNewLine & "Click OK Button To Return.", vbOKOnly + vbCritical, gsLocalForm
        cmdDeleteDate.Enabled = False
      Else 'NOT DAFROMDATE...
        Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay WHERE ((CodeDay.CODE_DATE) Between " _
                        & "#" & Format$(daFromDate, "mm/dd/yy") & "#" & " And " _
                        & "#" & Format$(daToDate, "mm/dd/yy") & "#" & ")")
        With rsCodeDay
            .MoveFirst
            .MoveLast
            lCountRec = .RecordCount
            .MoveFirst
            For lCountRec = 1 To lCountRec
                .Delete
                .MoveNext
            Next lCountRec
            .Close
        End With 'RSCODEDAY
        Set rsCodeDay = Nothing
        cmdCancelDate_Click
        MsgBox (lCountRec - 1) & " Records have been deleted!" _
                & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If

End Sub

Private Sub cmdDeleteTitle_Click()

  Dim rsCodeDay As Recordset

    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    With rsCodeDay
        If .RecordCount > 0 Then
            .FindFirst "CODE_TITLE = '" & txtTitle.Text & "'"
            If Not .NoMatch Then
                .Delete
                .Close
            End If
            Call_DoMBBeenDel
        End If
    End With 'RSCODEDAY
    Set rsCodeDay = Nothing
    lstTitle.Clear
    cmdDeleteTitle.Enabled = False
    cmdDeleteByDate.Enabled = False
    Call_DoBoxClear
    Call_DoEmptyTable

End Sub

Private Sub cmdDisplay_Click()

  Dim rsCodeDay As Recordset
  Dim lCountRec As Long

    cmdImport.Enabled = False
    lblEmails.Caption = "0 Emails"
    plNumberOfTitles = 0
    Screen.MousePointer = vbHourglass
    lstTitle.Clear
    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    With rsCodeDay
        If .RecordCount > 0 Then
            .MoveFirst
            .MoveLast
            lCountRec = .RecordCount
            .MoveFirst
            For lCountRec = 1 To lCountRec
                lstTitle.AddItem .Fields("CODE_TITLE")
                plNumberOfTitles = plNumberOfTitles + 1
                lblTitle.Caption = plNumberOfTitles & " Titles"
                DoEvents
                .MoveNext
            Next lCountRec
          Else 'NOT .RECORDCOUNT...
            Call_DoEmptyTable
        End If
        .Close
    End With 'RSCODEDAY
    Screen.MousePointer = vbDefault
    Set rsCodeDay = Nothing
    cmdImport.Enabled = True

End Sub

Private Sub cmdImport_Click()

  Dim iNull As Long
  Dim lpIDList As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo
  Dim sFileName As String
  Dim lCountEmlFiles As Long

    cmdImport.Enabled = False
    Call_DoBoxClear
    lblTitle.Caption = "0 Title"
    lblEmails.Caption = "0 Emails"
    lstTitle.Clear
    cmdDisplay.Enabled = False
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
    End If
    sFileName = Dir(Trim$(sPath) & "*.eml", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    Do While Len(sFileName) <> 0
        sFileName = Dir()  ' Get next file.
        lCountEmlFiles = lCountEmlFiles + 1
    Loop
    If lCountEmlFiles > 0 Then
        Call_DoImportEmails sPath
      Else 'NOT LCOUNTEMLFILES...
        MsgBox "No Email files found on   " & sPath _
               & vbNewLine & "Please select other directory." _
               & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm
    End If
    Call_DoEmptyTable
    cmdImport.Enabled = True
    cmdDisplay.Enabled = True

End Sub

Private Sub cmdMainMenu_Click()

    Call_ThisFormSize
    frmStartMenu.Show
    Unload Me

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
    Call_DoEmptyTable
    cmdDeleteTitle.Enabled = False
    cmdDeleteByDate.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

Private Function Func_GetDate(ByVal sGetDate As String) As Date

  Dim lCountChar As Long

    lCountChar = Len(sGetDate)
    For lCountChar = 1 To lCountChar
        If Mid$(sGetDate, lCountChar, 1) = "," Then
            sGetDate = Right$(sGetDate, (Len(sGetDate) - lCountChar - 1))
            Exit For 'loop varying lcountchar
        End If
    Next lCountChar
    Func_GetDate = Format$(sGetDate, "mm/dd/yy")

End Function

Private Function Func_GetTitle(ByVal sGetTitle As String) As String

  Dim lCountChar As Long

    lCountChar = Len(sGetTitle)
    For lCountChar = 1 To lCountChar
        If Mid$(sGetTitle, lCountChar, 1) = ")" Then
            sGetTitle = Right$(sGetTitle, (Len(sGetTitle) - lCountChar)) '- 1
            Exit For 'loop varying lcountchar
        End If
    Next lCountChar
    Func_GetTitle = Trim$(sGetTitle)

End Function

Private Function Func_IsSpace(ByVal sIsSpace As String) As Boolean

  Dim lCountChar As Long

    lCountChar = Len(sIsSpace)
    For lCountChar = 1 To lCountChar
        If Mid$(sIsSpace, lCountChar, 1) = ")" Then
            sIsSpace = Mid$(sIsSpace, lCountChar + 1, 1)
            If sIsSpace <> " " Then
                Func_IsSpace = True
            End If
            Exit For 'loop varying lcountchar
        End If
    Next lCountChar

End Function

Private Function Func_WeekDays(ByVal sWeekDay As String) As Boolean

  Dim lCountChar As Long

    lCountChar = Len(sWeekDay)
    For lCountChar = 1 To lCountChar
        If Mid$(sWeekDay, lCountChar, 1) = "," Then
            sWeekDay = Left$(sWeekDay, lCountChar - 1)
            Exit For 'loop varying lcountchar
        End If
    Next lCountChar
    Select Case True
      Case sWeekDay = "Sunday"
        Func_WeekDays = True
      Case sWeekDay = "Monday"
        Func_WeekDays = True
      Case sWeekDay = "Tuesday"
        Func_WeekDays = True
      Case sWeekDay = "Wednesday"
        Func_WeekDays = True
      Case sWeekDay = "Thursday"
        Func_WeekDays = True
      Case sWeekDay = "Friday"
        Func_WeekDays = True
      Case sWeekDay = "Saturday"
        Func_WeekDays = True
      Case Else
        Func_WeekDays = False
    End Select

End Function

Private Sub lstTitle_Click()

  Dim rsCodeDay As Recordset

    Call_DoBoxClear
    Set rsCodeDay = DB1.OpenRecordset("SELECT * FROM CodeDay")
    With rsCodeDay
        If .RecordCount > 0 Then
            .FindFirst "CODE_TITLE = '" & lstTitle.Text & "'"
            If Not .NoMatch Then
                txtDate.Text = Format$(.Fields("CODE_DATE"), "dd mmmm yyyy")
                txtCategory.Text = .Fields("CODE_CATEGORY")
                txtTitle.Text = .Fields("CODE_TITLE")
                txtWebAddress.Text = .Fields("CODE_WEB_ADDRESS")
            End If
            .Close
            cmdDeleteTitle.Enabled = True
            cmdDeleteByDate.Enabled = True
          Else 'NOT .RECORDCOUNT...
            cmdDeleteTitle.Enabled = False
            cmdDeleteByDate.Enabled = False
        End If
    End With 'RSCODEDAY
    Set rsCodeDay = Nothing

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

Private Sub mnuImportZipFilesItem_Click()

    Call_ThisFormSize
    frmImportFiles.Show
    Unload Me

End Sub

Private Sub mnuScreenDefaultPositionItem_Click()

    frmScreenDefault.Show
    Unload Me

End Sub

Private Sub txtFromDate_Change()

    If IsDate(txtFromDate.Text) Then
        If IsDate(txtToDate.Text) Then
            cmdDeleteDate.Enabled = True
        End If
    End If

End Sub

Private Sub txtFromDate_Click()

    plCalendar = 1
    Calendar1.Top = 0
    With Calendar1
        .Year = Year(Date)
        .Month = Month(Date)
        .Day = Day(Date)
    End With 'CALENDAR1
    Calendar1.SetFocus

End Sub

Private Sub txtToDate_Change()

    If IsDate(txtFromDate.Text) Then
        If IsDate(txtToDate.Text) Then
            cmdDeleteDate.Enabled = True
        End If
    End If

End Sub

Private Sub txtToDate_Click()

    plCalendar = 2
    Calendar1.Top = 0 '240
    With Calendar1
        .Year = Year(Date)
        .Month = Month(Date)
        .Day = Day(Date)
    End With 'CALENDAR1
    Calendar1.SetFocus

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:38) 5 + 631 = 636 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-16 16:47)  Decl: 5  Code: 645  Total: 650 Lines
':) CommentOnly: 16 (2.5%)  Commented: 33 (5.1%)  Empty: 122 (18.8%)  Max Logic Depth: 10
