VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonControlStrings"
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

''' <summary>A local (convenience) implementation of {IControlStrings}.</summary
''' <remarks>
''' The implementation published by {PGSolutions.RibbonDispatcher} could have been used
''' - but that introduces an undesirable coupling as well as redundant martialing of the
''' component text strings on start-up.
''' </remarks>
Private Type TControlStrings
    Label          As String
    ScreenTip      As String
    SuperTip       As String
    KeyTip         As String
    Description    As String
End Type
Private This As TControlStrings

Implements PGSolutions_RibbonDispatcher.IControlStrings
Implements PGSolutions_RibbonDispatcher.IControlStrings2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IControlStrings_KeyTip() As String
    IControlStrings_KeyTip = This.KeyTip
End Property

Private Property Get IControlStrings_Label() As String
    IControlStrings_Label = This.Label
End Property

Private Property Get IControlStrings_ScreenTip() As String
    IControlStrings_ScreenTip = This.ScreenTip
End Property

Private Property Get IControlStrings_SuperTip() As String
    IControlStrings_SuperTip = This.SuperTip
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IControlStrings2_Description() As String
    IControlStrings2_Description = This.Description
End Property

Private Property Get IControlStrings2_KeyTip() As String
    IControlStrings2_KeyTip = This.KeyTip
End Property

Private Property Get IControlStrings2_Label() As String
    IControlStrings2_Label = This.Label
End Property

Private Property Get IControlStrings2_ScreenTip() As String
    IControlStrings2_ScreenTip = This.ScreenTip
End Property

Private Property Get IControlStrings2_SuperTip() As String
    IControlStrings2_SuperTip = This.SuperTip
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Initialize(ByVal Label As String, _
    Optional ByVal ScreenTip As String, _
    Optional ByVal SuperTip As String, _
    Optional ByVal KeyTip As String _
) As IControlStrings
    This.Label = Label
    This.ScreenTip = ScreenTip
    This.SuperTip = SuperTip
    This.KeyTip = KeyTip

    Set Initialize = Me
End Function

Public Function Initialize2(ByVal Label As String, _
    Optional ByVal ScreenTip As String, _
    Optional ByVal SuperTip As String, _
    Optional ByVal KeyTip As String, _
    Optional ByVal Description As String _
) As IControlStrings2
    This.Label = Label
    This.ScreenTip = ScreenTip
    This.SuperTip = SuperTip
    This.KeyTip = KeyTip
    This.Description = Description

    Set Initialize2 = Me
End Function
