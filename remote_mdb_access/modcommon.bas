Attribute VB_Name = "modcommon"
Option Explicit

Public Function extractstring(ByVal str$, ByVal cmp$, ByVal no%) As String
Dim arr() As String
arr = Split(str, cmp)
If no <= UBound(arr) Then
    extractstring = arr(no)
Else
    extractstring = ""
End If

End Function

Public Function addstrap(ByVal path1 As String, ByVal path2 As String) As String
If Right$(path1, 1) = "\" Then
     addstrap = path1 & path2
Else
         addstrap = path1 & "\" & path2
End If
End Function

