Attribute VB_Name = "PublicFunctions"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit

Public Function Env(Value As Variant) As String
    Env = Environ(Value)
End Function

Public Function DeskTop(Optional ByVal AllUsers As Boolean = False) As String
    DeskTop = IIf(AllUsers, _
            CreateObject("WScript.Shell").SpecialFolders("AllUsersDesktop"), _
            CreateObject("WScript.Shell").SpecialFolders("Desktop")) _
            & "\"
End Function
