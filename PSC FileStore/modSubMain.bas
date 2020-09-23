Attribute VB_Name = "modSubMain"
Option Explicit

Public Sub Main()

    gsProgName = "PSC FileStore"
    gsOwner = "Planet Source Code"
    gsProgVer = "Version " & App.Major & "." & App.Minor & " Build " & App.Revision
    If App.PrevInstance Then
        Call_DoIsRunning
        End
    End If
    Set DB1 = OpenDatabase(App.Path & "\PSCFileStore.mdb", False, False, ";pwd=")
    frmStartMenu.Show

End Sub

':)Code Fixer V3.0.9 (04/08/2005 18:02:24) 1 + 17 = 18 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 19  Total: 20 Lines
':) CommentOnly: 1 (5%)  Commented: 0 (0%)  Empty: 4 (20%)  Max Logic Depth: 2
