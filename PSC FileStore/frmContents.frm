VERSION 5.00
Begin VB.Form frmContents 
   Caption         =   "Contents"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
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
      Left            =   2940
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtReadme 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  Dim FSys As New FileSystemObject
  Dim tsFile As TextStream
  Dim sReadme As String

    Me.Caption = gsProgName & " - " & Me.Caption & " - " & gsOwner
    Me.Move ((Screen.Width - Me.Width) \ 2), ((Screen.Height - Me.Height) \ 2)
    Set tsFile = FSys.OpenTextFile(App.Path & "\ReadmePlease.txt", ForReading, False)
    With tsFile
        Do While .AtEndOfStream <> True
            sReadme = sReadme & .ReadLine & vbNewLine
        Loop
        .Close
    End With 'TSFILE
    txtReadme.Text = sReadme

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmContents = Nothing

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:40) 1 + 34 = 35 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 36  Total: 37 Lines
':) CommentOnly: 1 (2.7%)  Commented: 1 (2.7%)  Empty: 11 (29.7%)  Max Logic Depth: 3
