Attribute VB_Name = "modFRM"
Option Explicit

Public Sub Call_DoCategoryNoFiles(ByVal sMessage As String)

    MsgBox "Category " & sMessage & " has NO files!" _
           & vbNewLine & "Please try other Category." _
           & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Public Sub Call_DoMBBeenDel()

    MsgBox "The Record Has Been Deleted!" _
           & vbNewLine & "Click OK Button To Return.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Public Sub Call_DoMBDatabaseEmpty()

    MsgBox "Database is Empty!" _
           & vbNewLine & "Please Click The ADD Button To Begin Entering Data." _
           & vbNewLine & "Click OK Button To Continue.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Public Sub Call_DoMBEditUpdate()

    MsgBox "The Record Has Been Updated!" _
           & vbNewLine & "Click OK Button To Return.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Public Sub Call_DoMBNewRecAdded()

    MsgBox "The New Record Has Been Added!" _
           & vbNewLine & "Click OK Button To Return.", vbOKOnly + vbInformation, gsLocalForm

End Sub

Public Function Func_DoMBAddNewRec() As String

    Func_DoMBAddNewRec = MsgBox("Do you want to Add a New Record?" _
                         & vbNewLine & "Click YES to save it or NO to Cancel.", vbYesNo + vbQuestion, gsLocalForm)

End Function

Public Function Func_DoMBNewEditRec() As String

    Func_DoMBNewEditRec = MsgBox("Do you want to change this Record?" _
                          & vbNewLine & "Click YES to save it or NO to Cancel.", vbYesNo + vbQuestion, gsLocalForm)

End Function

Public Function Func_DoMBPositiveDel() As String

    Func_DoMBPositiveDel = MsgBox("Are You Positive you want to Delete This Record?" _
                           & vbNewLine & "Click YES to Delete it or NO to Cancel.", _
                           vbYesNo + vbQuestion + vbDefaultButton2, gsLocalForm)

End Function

Public Function Func_DoMBSureDel(ByVal sMessage As String) As String

    Func_DoMBSureDel = MsgBox("Are You Sure you want to Delete This Record?" _
                       & vbNewLine & sMessage & " will be erased and CANNOT be recovered!" _
                       & vbNewLine & "Click YES to Delete it or NO to Cancel.", _
                       vbYesNo + vbQuestion + vbDefaultButton2, gsLocalForm)

End Function

':)Code Fixer V3.0.9 (04/08/2005 18:02:33) 1 + 57 = 58 Lines Thanks Ulli for inspiration and lots of code.

':) Ulli's VB Code Formatter V2.17.9 (2005-Aug-09 21:50)  Decl: 1  Code: 73  Total: 74 Lines
':) CommentOnly: 1 (1.4%)  Commented: 0 (0%)  Empty: 28 (37.8%)  Max Logic Depth: 1
