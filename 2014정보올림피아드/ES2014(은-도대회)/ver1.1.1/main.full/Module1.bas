Attribute VB_Name = "Module1"
Option Explicit
Public Function DirTextFile() As Boolean
DirTextFile = False
If Dir$(App.Path & "\Save\English.ini") <> "" Then
    DirTextFile = True
End If
End Function
Public Function DirResultFile() As Boolean
DirResultFile = False
If Dir$(App.Path & "\Result\Result.ini") <> "" Then
    DirResultFile = True
End If
End Function

