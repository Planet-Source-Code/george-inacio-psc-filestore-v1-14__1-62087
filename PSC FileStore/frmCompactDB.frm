VERSION 5.00
Begin VB.Form frmCompactDB 
   Caption         =   "Compact and repair Database"
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleMode       =   0  'User
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompact 
      Caption         =   "Compact"
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
      Left            =   5569
      TabIndex        =   4
      Top             =   5500
      Width           =   1500
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
      Left            =   4076
      TabIndex        =   0
      ToolTipText     =   "Returns to Main Menu."
      Top             =   5500
      Width           =   1500
   End
   Begin VB.Label lblCompactDone 
      Alignment       =   2  'Center
      Caption         =   "Compacting and repair the Database has been done!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   500
      Left            =   135
      TabIndex        =   5
      Top             =   4080
      Width           =   10875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Compact and repair Database"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   9735
   End
   Begin VB.Label Label2 
      Caption         =   "Click the button to Compact and repair the Database."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1572
      TabIndex        =   2
      Top             =   3216
      Width           =   8000
   End
   Begin VB.Label Label3 
      Caption         =   "If you are being doing a lot of deleting the Database will reduce is size."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1572
      TabIndex        =   1
      Top             =   3536
      Width           =   8000
   End
End
Attribute VB_Name = "frmCompactDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Call_DoGoneOut()

  'Saving to the register the size and position of form before leaving the form

    With Me
        SaveSetting "PSC Soft", "PSC FileStore", "Height", .Height
        SaveSetting "PSC Soft", "PSC FileStore", "Left", .Left
        SaveSetting "PSC Soft", "PSC FileStore", "Top", .Top
        SaveSetting "PSC Soft", "PSC FileStore", "Width", .Width
    End With 'Me
    Set frmCompactDB = Nothing

End Sub

Private Sub Call_ThisFormSize()

    With Me
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

End Sub

Private Sub cmdCompact_Click()

  Dim FSys As New FileSystemObject

    Screen.MousePointer = vbHourglass
    DB1.Close
    Set DB1 = Nothing
    Name App.Path & "\PSCFileStore.mdb" As App.Path & "\PSCFileStoreOld.mdb"
    DBEngine.CompactDatabase App.Path & "\PSCFileStoreOld.mdb", App.Path & "\PSCFileStore.mdb"
    If FSys.FileExists(App.Path & "\PSCFileStoreOld.mdb") Then
        FSys.DeleteFile (App.Path & "\PSCFileStoreOld.mdb")
    End If
    Set DB1 = OpenDatabase(App.Path & "\PSCFileStore.mdb", False, False, ";pwd=")
    lblCompactDone.Caption = "Compacting and repair the Database has been done!"
    cmdCompact.Enabled = False
    Screen.MousePointer = vbDefault

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
    lblCompactDone.Caption = vbNullString

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:36) 1 + 76 = 77 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 77  Total: 78 Lines
':) CommentOnly: 2 (2.6%)  Commented: 3 (3.8%)  Empty: 21 (26.9%)  Max Logic Depth: 2
