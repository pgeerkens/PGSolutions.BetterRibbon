Attribute VB_Name = "PublicFunctions"
Option Explicit

Public Function Env(Value As Variant) As String
    Env = Environ(Value)
End Function

Public Function DeskTop() As String
    DeskTop = CreateObject("WScript.Shell").specialfolders("Desktop")
End Function
