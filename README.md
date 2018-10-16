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
run time, is as simple as this, in two modules:

**RibbonModel:**

    Option Explicit

    Private WithEvents Button1  As RibbonDispatcherX.RibbonButton
    Private Button1Strings      As RibbonDispatcherX.IRibbonControlStrings

    Private Sub Button1_Clicked()
        MsgBox "Button1 clicked.", vbOKOnly Or vbInformation, TypeName(Me)
    End Sub

    Friend Sub Activate()
        With AddInHandle
            Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
            Button1.SetImageMso "RefreshAll"
        End With
    End Sub

    Private Sub Class_Initialize()
        With AddInHandle
            Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                    ScreenTip:="Button1 Screentip", _
                    AuperTip:="Lots of good things" & vbNewLine & _
                              "can be done here to" & vbNewLine & _
                              "show off a bit.", keyTip:="", _
                    alternateLabel:="", Description:="")
        End With
    End Sub

    Private Function AddInHandle() As RibbonDispatcherX.IRibbonDispatcher
        Set AddInHandle = Application.COMAddIns("ExcelRibbon").Object
    End Function
    
and **ThisWorkbook:**

    Option Explicit
    Private MRibbonModel    As SingleButton.RibbonModel

    Private Sub Workbook_Activate()
        MRibbonModel.Activate
    End Sub

    Private Sub Workbook_Open()
        Set MRibbonModel = New RibbonModel
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
