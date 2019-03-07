# PGSolutions.BetterRibbon
**The Fluent Ribbon - done right.  A generic dispatcher greatly easing Ribbon development in both .NET &amp; VBA**

Tied up in knots programming the Fluent Ribbon? Wondering why you bought an MS-Office with no wiring at the front desk? Here is
an example of how the Fluent Ribbon was **meant** to be delivered.

This solution uses generic callbacks eponymous with their XM tags, with a dictionary lookup on each callback invocation, to
automatically route each to its appropriate control. This implementation provides not just a pre-programmed ViewModel for all one's
Ribbon controls, but a Model template as well. 

The supplied Add-In also comes with a built in suite of Ribbon controls that can be dynamically attached to:
- 1 ToggleButton;
- 3 CheckBoxes;
- 3 DropDowns; and
- 3 (Action)Buttons.

The code to activate one of the (Action)Buttons, with Click event handler programmed in VBA and custom text defined at
run time, is as simple as this, in three modules:

**RibbonModel**, declaring, initializing and handling the events from the controls used:

    Option Explicit
    Private CustomGroup         As IGroupModel
    Private WithEvents Button1  As ButtonModel
    Attribute Button1.VB_VarHelpID = -1

    Private Sub Button1_Clicked(ByVal control As IRibbonControl)
        ButtonProcessing.Button1_Processing
    End Sub

    Private Sub Class_Initialize()
        Dim Strings As IControlStrings
        With ThisWorkbook.ModelServer
            Set CustomGroup = .GetGroupModel("CustomizableGroup")

            Set Button1 = .GetButtonModel("CustomizableButton1") _
                          .SetImage(.NewImageObjectMso("MacroPlay"))
        End With
    End Sub
    
**ResourceLoader**, supplying string and image resources to the dispatcher:

    Option Explicit
    Implements IResourceLoader

    ''' <summary>Serving Button, ToggleButton, CheckBox, Menu, and Gallery controls. </summary>
    Private Function IResourceLoader_GetControlStrings2(ByVal ControlId As String) As IControlStrings2
        Select Case (ControlId)
            Case "CustomizableButton1":
                Set IResourceLoader_GetControlStrings2 = ThisWorkbook.ModelServer.NewControlStrings2( _
                    Label:="This is cool!", _
                    ScreenTip:="VBA-Customized Button Screentip", _
                    SuperTip:="This button is completely" & vbNewLine & _
                              "customized within the VBA" & vbNewLine & _
                              "workbook.", KeyTip:="", Description:="")
            Case Else:
                Set IResourceLoader_GetControlStrings2 = Nothing
        End Select
    End Function

    ''' <summary>Serving all other controls. </summary>
    Private Function IResourceLoader_GetControlStrings(ByVal ControlId As String) As IControlStrings
        Select Case (ControlId)
            Case "CustomizableGroup":
                Set IResourceLoader_GetControlStrings = ThisWorkbook.ModelServer.NewControlStrings( _
                    Label:="VBA Custom Controls", _
                    ScreenTip:="", SuperTip:="", KeyTip:="")
            Case Else:
                Set IResourceLoader_GetControlStrings = Nothing
        End Select
    End Function

    Private Function IResourceLoader_GetImage(ByVal Name As String) As Variant
        IResourceLoader_GetImage = "MacroSecurity"
    End Function

    
and **ThisWorkbook**, initializng the connection to the dispatcher.

    Option Explicit
    Private Const COMAddInName  As String = "PGSolutions.BetterRibbon"
    Private MModelServer        As IModelServer
    Private MRibbonModel        As RibbonModel

    Public Function ModelServer() As IModelServer
        If MModelServer Is Nothing Then
            Set MModelServer = Application.COMAddIns(COMAddInName).Object _
                    .NewBetterRibbon(New ResourceLoader)
        End If
        Set ModelServer = MModelServer
    End Function

    Private Sub Workbook_Activate()
        If MRibbonModel Is Nothing Then
            Application.COMAddIns(COMAddInName).Object.RegisterWorkbook ThisWorkbook.Name
            Set MRibbonModel = New RibbonModel
        End If
    End Sub

    ' Depending on Workbook location, pops-up a dialog to assist with setting debug breakpoints.
    Private Sub Workbook_Open()
        If DeskTop(False) = "D:\Users\Pieter\Desktop\" _
        Or ThisWorkbook.Path = DeskTop(True) & "Example Workbooks" _
        Or ThisWorkbook.Path = DeskTop(False) & "Example Workbooks" Then _
            MsgBox "Pause for Ctrl-Break to ease debugging." & vbNewLine & vbNewLine & _
                   "This message can be disabled by moving the workbook" & vbNewLine & _
                   "out of the Desktop folder 'Example Workbooks'.", _
                   vbOKOnly, ThisWorkbook.Name
    End Sub

The *Workbook_Activate* event is programmed as the controls for this workbook are automatically deactivated when the workbook loses
focus. So the preprogrammed customizable controls from the Add-In are all available to every workbook.

---

Also included:

1. VBA Source Exporter that unloads all VBA code for an MS-Excel or MS-Access project to a directory sibling to the workbook/database;
 named either eponymously with the suffix VBA or *.\src*. A great time saver for managing VBA code in source control. There is a Ribbon Group in the Add-In with three controls accessing this functionality:

 - A **Toggle** between use of the directory *'\src* and a directory eponymouswith the workbook/database name.
 - A **Selected Projects** (Action)Button for exporting VBA code from a directory of projects.
 - A **Current Workbook** (Action)Button that immediately exports the VBA code from the current workbook.

1. External Links Analyzer that collects details on all External Links in a list of Workbooks, and reports them to three worksheets in
the current worlkbook: *Links Analysis*, *Externally Linked Files List*, and *Parsing Failures*.
