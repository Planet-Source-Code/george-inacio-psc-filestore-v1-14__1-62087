VERSION 5.00
Begin VB.Form frmScreenDefault 
   Caption         =   "Screen Default Settings"
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
      Left            =   4822
      TabIndex        =   0
      ToolTipText     =   "Returns to Main Menu."
      Top             =   5500
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Screen Default Settings"
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
      TabIndex        =   6
      Top             =   480
      Width           =   9735
   End
   Begin VB.Label Label2 
      Caption         =   "Program window position and size have been set to the default."
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
      TabIndex        =   5
      Top             =   2542
      Width           =   8000
   End
   Begin VB.Label Label3 
      Caption         =   "To set the position and window size to your liking, do the following:"
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
      TabIndex        =   4
      Top             =   2862
      Width           =   8000
   End
   Begin VB.Label Label4 
      Caption         =   "1. Drag the window to where you want the starting point;"
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
      TabIndex        =   3
      Top             =   3182
      Width           =   8000
   End
   Begin VB.Label Label5 
      Caption         =   "2. Resize the window to the size you like or maximise it (needs resizer)."
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
      Top             =   3502
      Width           =   8000
   End
   Begin VB.Label Label10 
      Caption         =   "Please click the OK button to exit this window!"
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
      Top             =   4210
      Width           =   8000
   End
End
Attribute VB_Name = "frmScreenDefault"
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
    Set frmScreenDefault = Nothing

End Sub

Private Sub Call_ThisFormSize()

    With Me
        glFormHeight = .Height
        glFormLeft = .Left
        glFormTop = .Top
        glFormWidth = .Width
    End With 'ME

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
        .Move ((Screen.Width - .Width) \ 2), ((Screen.Height - .Height) \ 2)
        .Height = 7815
        .Width = 11265
    End With 'Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call_DoGoneOut

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:36) 1 + 54 = 55 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 56  Total: 57 Lines
':) CommentOnly: 2 (3.5%)  Commented: 3 (5.3%)  Empty: 17 (29.8%)  Max Logic Depth: 2
