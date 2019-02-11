VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceCellRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit

Private Type TSourceCellRef
    FullPath     As String
    FileName     As String
    TabName      As String
    CellName     As String
    IsNamedRange As Boolean
End Type
Private This As TSourceCellRef

Implements ISourceCellRef

Private Property Get ISourceCellRef_FullPath() As String
    ISourceCellRef_FullPath = This.FullPath
End Property
Private Property Get ISourceCellRef_FileName() As String
    ISourceCellRef_FileName = This.FileName
End Property
Private Property Get ISourceCellRef_TabName() As String
    ISourceCellRef_TabName = This.TabName
End Property
Private Property Get ISourceCellRef_CellName() As String
    ISourceCellRef_CellName = This.CellName
End Property
Private Property Get ISourceCellRef_IsNamedRange() As Boolean
    ISourceCellRef_IsNamedRange = This.IsNamedRange
End Property

Public Function Initialize( _
     ByVal FullPath As String, _
     ByVal FileName As String, _
     ByVal TabName As String, _
     ByVal CellName As String, _
     Optional ByVal IsNamedRange = False _
) As SourceCellRef
    This.FullPath = FullPath
    This.FileName = FileName
    This.TabName = TabName
    This.CellName = CellName
    This.IsNamedRange = IsNamedRange
End Function

Private Sub Class_Initialize()
    This.FullPath = "Full Path"
    This.FileName = "File name"
    This.TabName = "Tab Name"
    This.CellName = "$A$1"
    This.IsNamedRange = False
End Sub
