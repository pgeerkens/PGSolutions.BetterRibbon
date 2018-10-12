Attribute VB_Name = "RibbonUtils"
Option Explicit
Option Private Module
Private Const ModuleName As String = "RibbonUtilities."

Public Function RibbonFactory() As RibbonDispatcherX.IRibbonFactory
    On Error GoTo EH
    Set RibbonFactory = AddInHandle.RibbonFactory
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "RibbonFactory"
    Resume          ' for debugging only
End Function

Public Function AddInHandle() As RibbonDispatcherX.IMain
    On Error GoTo EH
    Set AddInHandle = Application.COMAddIns("ExcelRibbon").Object
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "AddInHandle"
    Resume          ' for debugging only
End Function
