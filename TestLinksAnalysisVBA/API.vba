Attribute VB_Name = "API"
Option Explicit

Private Const mModuleName   As String = "API."

Public Function FastCopyToRange(ByVal Source As ITwoDimensionalLookup, ByVal Destination As Excel.Range) As Excel.Range
    On Error GoTo EH
    Dim RowsCount As Long: RowsCount = Destination.Rows.Count
    Dim ColsCount As Long: ColsCount = Destination.Columns.Count
    Dim Data() As Variant: ReDim Data(0 To RowsCount - 1, 0 To ColsCount - 1) As Variant
    
    Dim RowNo As Long, ColNo As Long
    For RowNo = 0 To RowsCount - 1
        For ColNo = 0 To ColsCount - 1
            Data(RowNo, ColNo) = Source.Item(RowNo, ColNo)
        Next ColNo
    Next RowNo
    
    Destination.Value = Data
    
    Set FastCopyToRange = Destination
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, mModuleName & "FastCopyToRange"
    Resume
End Function


''' <summary>Performs Links Analysis on the suppplied workbook, returning an ExternalLinks collection.</summary>
Public Function LinksAnalysis( _
    ByVal wb As Excel.Workbook, _
    Optional ByVal ExcludedWorksheetName As String = "" _
) As IExternalLinks
Attribute LinksAnalysis.VB_Description = "Returns an ExternalLinks collection obtained by analyzing NamedRanes and CellFormulas in the supplied workbook for external links."
    On Error GoTo EH
    Application.ScreenUpdating = False
    
    Set LinksAnalysis = AddInHandle.NewExternalLinksWB(wb, ExcludedWorksheetName)
    
XT: Application.ScreenUpdating = True
    Exit Function
    
EH: ErrorUtils.ReRaiseError Err, mModuleName & "LinksAnalysis"
    Resume XT
    Resume      ' for debugging only
End Function
