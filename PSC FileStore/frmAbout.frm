VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Box"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5670
   ClipControls    =   0   'False
   HelpContextID   =   280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2205
      TabIndex        =   0
      Top             =   1785
      Width           =   1260
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   4
      Top             =   720
      Width           =   5085
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Program Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   293
      TabIndex        =   3
      Top             =   120
      Width           =   5085
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Alpha Tester: Miguel Inacio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   1425
      Width           =   5085
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Author: George Inacio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   300
      TabIndex        =   1
      Top             =   1065
      Width           =   5085
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Caption = gsProgName & " - " & Me.Caption & " - " & gsOwner
    Me.Move ((Screen.Width - Me.Width) \ 2), ((Screen.Height - Me.Height) \ 2)
    Label5.Caption = gsProgName
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & " Build " & App.Revision

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmAbout = Nothing

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:40) 1 + 24 = 25 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 26  Total: 27 Lines
':) CommentOnly: 1 (3.7%)  Commented: 0 (0%)  Empty: 10 (37%)  Max Logic Depth: 1
