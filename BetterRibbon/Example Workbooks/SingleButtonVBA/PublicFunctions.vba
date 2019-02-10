Attribute VB_Name = "PublicFunctions"
Option Explicit

Public Function Env(Value As Variant) As String
    Env = Environ(Value)
End Function

Public Function DeskTop(Optional ByVal AllUsers As Boolean = False) As String
    DeskTop = IIf(AllUsers, _
            CreateObject("WScript.Shell").SpecialFolders("AllUsersDesktop"), _
            CreateObject("WScript.Shell").SpecialFolders("Desktop"))
End Function
