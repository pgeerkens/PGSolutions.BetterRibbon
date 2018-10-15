Attribute VB_Name = "LinksManager"
Option Explicit
Option Private Module

Private Const mModuleName   As String = "LinksManager."
Private Const TargetName    As String = "Link Analysis"

''' <summary>TODO</summary>
Public Sub ListExternalLinksActiveWkbk( _
    ByVal ShowErrors As Boolean, _
    ByVal IncludeHyperlinks As Boolean _
)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim ProtectStructure As Boolean: ProtectStructure = ActiveWorkbook.ProtectStructure
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Protect , False
    
    DeleteOldAnalysisWorksheets True
            
    Dim Links As IExternalLinks
    Set Links = AddInHandle.NewExternalLinksWB(ActiveWorkbook, TargetWorksheetNameThis)
    With Links
        If .Count = 0 Then
            MsgBox "NoLinksFound", vbOKOnly Or vbInformation, "ListExternalLinksActiveWkbk"
        Else
            WriteFileList .Files, TargetWorksheetNameFilesFromThis
            WriteLinksAnalysis Links, TargetWorksheetNameThis, IncludeHyperlinks
            If ShowErrors And .Errors.Count > 0 Then WriteErrors TargetWorksheetNameErrors, .Errors
            
            On Error Resume Next
            ActiveWorkbook.Sheets(TargetWorksheetNameThis).Select
    
            On Error GoTo EH
        End If
    End With
            
XT: ActiveWorkbook.Protect , ProtectStructure
    Application.ScreenUpdating = True
    Exit Sub
    
EH: ErrorUtils.ReRaiseError Err, mModuleName & "ListExternalLinksActiveWkbk"
    Resume XT
    Resume      ' for debugging only
End Sub

''' <summary>TODO</summary>
Public Sub ListExternalLinksWkbkList( _
    ByVal rng As Range, _
    ByVal ShowErrors As Boolean, _
    ByVal IncludeHyperlinks As Boolean _
)
    On Error GoTo EH
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim ProtectStructure As Boolean: ProtectStructure = ActiveWorkbook.ProtectStructure
    If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Protect , False
    Application.Calculation = xlCalculationManual
            
    Dim row As Long
    Dim Files As Collection: Set Files = New Collection
    For row = 1 To rng.Rows.Count Step 1
        Dim sCell As String: sCell = rng.Cells(row, 1)
        
        On Error GoTo ErrHandlerDuplicateFile
        If sCell <> "" Then Files.Add sCell, sCell
        
        On Error GoTo EH
        DoEvents
    Next row
    
    DeleteOldAnalysisWorksheets False
        
    If Files.Count = 0 Then
        MsgBox "NoFilesFound", vbOKOnly Or vbInformation, "ListExternalLinksWkbkList"
    Else
        WriteFilesLinks Files, ShowErrors, IncludeHyperlinks
    End If
                       
XT: Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ActiveWorkbook.Protect , ProtectStructure
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
    
ErrHandlerDuplicateFile:
    If Err.Number = 457 Then Resume Next 'This key is already associated with an element of this collection

EH: ErrorUtils.ReRaiseError Err, mModuleName & "ListExternalLinksWkbkList"
    Resume XT
    Resume      ' for debugging only
End Sub

Private Property Get TargetWorksheetNameThis() As String
    TargetWorksheetNameThis = TargetName & " - This WkBk"
End Property
Private Property Get TargetWorksheetNameFilesFromThis() As String
    TargetWorksheetNameFilesFromThis = "Linked Files - This WkBk"
End Property

Private Property Get TargetWorksheetNameList() As String
    TargetWorksheetNameList = TargetName & " - WkBk List"
End Property
Private Property Get TargetWorksheetNameFilesFromList() As String
    TargetWorksheetNameFilesFromList = "Linked Files - WkBk List"
End Property

Private Property Get TargetWorksheetNameErrors() As String
    TargetWorksheetNameErrors = TargetName & " Errors"
End Property

''' <summary>TODO</summary>
Private Sub WriteFilesLinks( _
    ByVal Files As Collection, ByVal ShowErrors As Boolean, _
    ByVal IncludeHyperlinks As Boolean _
)
    On Error GoTo EH
    Dim Links As IExternalLinks
    Set Links = AddInHandle.NewExternalLinks(Application, Files)
    If Links.Count = 0 And Links.Errors.Count = 0 Then
        MsgBox "NoLinksFound", vbOKOnly Or vbInformation, "WriteFilesLinks"
    Else
        WriteFileList Links.ExternalFiles, TargetWorksheetNameFilesFromList
        WriteLinksAnalysis Links, TargetWorksheetNameList, IncludeHyperlinks
        If ShowErrors And Links.Errors.Count > 0 Then WriteErrors TargetWorksheetNameErrors, Links.Errors
        
        On Error Resume Next
        ActiveWorkbook.Sheets(TargetWorksheetNameList).Select

        On Error GoTo EH
    End If
      
XT: Exit Sub

EH: ReRaiseError Err, mModuleName & "WriteFilesLinks"
    Resume      ' for debugging only
End Sub

''' <summary>TODO</summary>
Public Sub WriteErrors(ByVal TargetName As String, ByVal List As IParseErrors)
    Const Message    As String = "Writing errors to worksheet ... "
    
    On Error GoTo EH
    Dim TargetWS        As Worksheet: Set TargetWS = CreateTargetWorksheet(ActiveWorkbook, TargetName)
    Dim LastRow         As Long:      LastRow = 2
    
    Dim pe As Variant, i As Long: i = 0
    For Each pe In List
        i = i + 1
        LastRow = LastRow + 1
        With pe
            Dim ws As Worksheet: Set ws = TargetWS
            Dim col As Long: col = 0
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .CellRef.Cell
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .CellRef.TabName
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .CellRef.FileName
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .Condition
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .CharPosition
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = "'" & .Formula
            col = col + 1: TargetWS.Cells(LastRow, col).Value2 = .CellRef.Path
        End With
        
        Dim Completion As Long: Completion = i * 100 / List.Count
        Application.StatusBar = Message & "(" & CStr(Completion) & "%)"
        DoEvents
    Next pe
    With LinksErrorColumns
        InitializeTargetWorksheet TargetWS, LastRow, LinksErrorColumns
    End With

XT: Exit Sub

EH: ReRaiseError Err, mModuleName & "WriteErrors"
    Resume      ' for debugging only
End Sub

Private Property Get LinksErrorColumns() As ColumnHeaders
    With New ColumnHeaders
        .Initialize _
            "Source" & vbNewLine & "Cell", _
            "Source" & vbNewLine & "Worksheet", _
            "Source" & vbNewLine & "FileName", _
            "Error" & vbNewLine & "Description", _
            "Char Pos", _
            "Source" & vbNewLine & "Formula", _
            "External" & vbNewLine & "Path"
        Set LinksErrorColumns = .This
    End With
End Property

Private Property Get LinksAnalysisColumns() As ColumnHeaders
    With New ColumnHeaders
        .Initialize _
            "Links Target", _
            "External" & vbNewLine & "Path", _
            "External" & vbNewLine & "FileName", _
            "External" & vbNewLine & "Worksheet", _
            "External" & vbNewLine & "Cell", _
            "Link" & vbNewLine & "Type", _
            "Links Source", _
            "Source" & vbNewLine & "Path", _
            "Source" & vbNewLine & "FileName", _
            "Source" & vbNewLine & "Worksheet", _
            "Source" & vbNewLine & "Cell", _
            "Source Formula"
        Set LinksAnalysisColumns = .This
    End With
End Property

''' <summary>TODO</summary>
''' <returns>Returns the number of the LastRow added to the target worksheet.</returns>
Private Sub WriteLinksAnalysis( _
    ByVal List As IExternalLinks, _
    ByVal TargetName As String, _
    Optional ByVal IncludeHyperlinks As Boolean = False _
)
    Const HyperLinksMessage As String = "Writing Hyperlinks: "
    
    On Error GoTo EH
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    If List.Count > 0 Then
        Dim List2D    As ITwoDimensionalLookup: Set List2D = List
        Dim MaxCol    As Long:          MaxCol = List2D.ColsCount
        
        Dim ws        As Worksheet: Set ws = ClearTargetWorksheet(ActiveWorkbook, TargetName)
        ws.Rows(MaxCol).EntireRow.NumberFormat = "Text"
        
        Dim FirstCell As Range:     Set FirstCell = ws.Cells(3, 1)
        Dim LastCell  As Range:     Set LastCell = ws.Cells(List2D.RowsCount + 2, MaxCol)
        Dim SheetData As Range:     Set SheetData = ws.Range(FirstCell, LastCell)
        
        API.FastCopyToRange List2D, SheetData
        
        If IncludeHyperlinks Then
            Dim i As Long, FileName As String, Range As String
            For i = 1 To List.Count
                With List.ItemByIndex(i)
                    WriteLinkHyperlink SheetData.Cells(i, 1), .Path, .FileName, .TabName, .Cell
                    If .LinkType = "Cell Reference" Then _
                        WriteLinkHyperlink SheetData.Cells(i, 7), .SourcePath & "\", .SourceFile, .SourceTab, .SourceCell
                End With
                
                If SheetData.Hyperlinks.Count > 65500 Then Exit For
                Dim Completion As Long: Completion = i * 100 / List.Count
                Application.StatusBar = HyperLinksMessage & "(" & CStr(Completion) & "%)"
                DoEvents
            Next i
        End If
        
        With LinksAnalysisColumns
            InitializeTargetWorksheet ws, LastCell.row, LinksAnalysisColumns
        End With
    End If
    
XT: Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
EH: Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ReRaiseError Err, mModuleName & "WriteLinksAnalysis"
    Resume      ' for debugging only
End Sub

Private Sub WriteLinkHyperlink(ByVal Anchor As Range, ByVal Path As String, ByVal FileName As String, ByVal TabName As String, CellAddress As String)
    Const HyperlinkInfo     As String = vbNewLine & vbNewLine & "Ctrl-LeftClick to select this cell;" & _
                                        vbNewLine & vbNewLine & "Left Click to follow link"
    On Error GoTo EH
    Dim Address         As String:     Address = Path & FileName
    Dim SubAddress      As String
    If CellAddress <> "" Then
        SubAddress = "'" & TabName & "'!" & IIf(CellAddress = "#REF!", "", CellAddress)
    Else
        SubAddress = ""
    End If
    If CellAddress <> "#REF!" And CellAddress <> "" Then SubAddress = Application.ConvertFormula(SubAddress, xlA1, xlA1, xlAbsolute)
    
    With Anchor.Worksheet.Hyperlinks
        On Error Resume Next    ' Limit of 65535 links or so - blast through anyways
        If .Count < 65535 Then .Add Anchor, "file://" & Address, SubAddress, _
                "file://" & Address & SubAddress & HyperlinkInfo, _
                Address & SubAddress
    End With
XT: Exit Sub

EH: ReRaiseError Err, mModuleName & "WriteLinkHyperlink"
    Resume      ' for debugging only
End Sub

Private Sub WriteFileList(ByVal Files As IExternalFiles, ByVal Target As String)
    Const Message    As String = "Writing links to worksheet ... "
    
    On Error GoTo EH
    With ClearTargetWorksheet(ActiveWorkbook, Target)
        Dim FileName As String, row As Long: row = 0
        Dim i As Long
        For i = 0 To Files.Count - 1
            FileName = Files.Item(i)
            row = row + 1: .Cells(row, 1).Value = FileName
            WriteLinkHyperlink .Cells(row, 1), "", CStr(FileName), "", ""
            Dim Completion As Long: Completion = row * 100 / Files.Count
            Application.StatusBar = "Writing File List: " & "(" & CStr(Completion) & "%)"
        Next i
        
'        For Each FileName In Files
'            row = row + 1: .Cells(row, 1).Value = FileName
'            WriteLinkHyperlink .Cells(row, 1), "", CStr(FileName), "", ""
'            Dim Completion As Long: Completion = row * 100 / Files.Count
'            Application.StatusBar = "Writing File List: " & "(" & CStr(Completion) & "%)"
'        Next FileName
        .Range("$A:$A").EntireColumn.AutoFit
    End With
    
XT: Exit Sub
    
EH: ReRaiseError Err, mModuleName & "WriteFileList"
    Resume      ' for debugging only
End Sub

''' <summary>TODO</summary>
Private Sub DeleteTargetWorksheet(ByVal Name As String)
    On Error GoTo EH
    Application.DisplayAlerts = False
    Sheets(Name).Delete
    Application.DisplayAlerts = True
XT: Exit Sub

EH: If Err.Number = 9 Then Resume Next
    ReRaiseError Err, mModuleName & "DeleteTargetWorksheet"
    Resume      ' for debugging only
End Sub

''' <summary>TODO</summary>
Private Function ClearTargetWorksheet( _
    ByVal wb As Workbook, _
    ByVal Name As String _
) As Excel.Worksheet
    On Error GoTo EH
    Set ClearTargetWorksheet = wb.Worksheets(Name)
    
    If Not ClearTargetWorksheet Is Nothing Then DeleteTargetWorksheet Name
    Set ClearTargetWorksheet = CreateTargetWorksheet(wb, Name)

XT: Exit Function

EH: If Err.Number = 9 Then Resume Next
    ReRaiseError Err, mModuleName & "ClearTargetWorksheet"
    Resume      ' for debugging only
End Function

''' <summary>TODO</summary>
Private Function CreateTargetWorksheet(ByVal wb As Workbook, ByVal Name As String) As Excel.Worksheet
    On Error GoTo EH
    DeleteTargetWorksheet Name

    Dim TargetWS As Worksheet: Set TargetWS = wb.Worksheets.Add(before:=wb.Worksheets(1))
    TargetWS.Name = Name
    
    Set CreateTargetWorksheet = TargetWS
XT: Exit Function

EH: ReRaiseError Err, mModuleName & "CreateTargetWorksheet"
    Resume      ' for debugging only
End Function

''' <summary>TODO</summary>
Private Sub InitializeTargetWorksheet(ByVal TargetWS As Excel.Worksheet, ByVal LastRow As Long, ByVal ColHeaders As ColumnHeaders)
    On Error GoTo EH
    With TargetWS
        .Range("A1:D1").Merge
        .Range("A1").Formula = "Link File run on " & VBA.Format(Now, "mm/dd/yyyy hh:mm")
        
        Dim ColNo As Long
        For ColNo = 1 To ColHeaders.Count
            .Cells(2, ColNo).Value = ColHeaders.ItemByIndex(ColNo)
            DoEvents
        Next ColNo
    End With
    
    FormatTargetWorksheet TargetWS ', LastRow
    
    TargetWS.Range("$A$2", TargetWS.Cells(LastRow, ColHeaders.Count)).AutoFilter
    ActiveWindow.FreezePanes = False
    TargetWS.Rows("$3:$3").EntireRow.Select
    ActiveWindow.FreezePanes = True
    
XT: Exit Sub

EH: ReRaiseError Err, mModuleName & "InitializeTargetWorksheet"
    Resume      ' for debugging only
End Sub
''' <summary>TODO</summary>
Private Sub FormatTargetWorksheet(ByVal TargetWS As Worksheet)
    On Error GoTo EH
    With TargetWS
        .Parent.Activate
        .Activate
        
        .Columns("$H:$J").WrapText = False
        
        .Range("$1:$2").Font.Bold = True
        .Rows("$2").RowHeight = 60
        .Range("A2").EntireRow.WrapText = True
        .Range("A2").EntireRow.HorizontalAlignment = xlHAlignCenter
        
        .Columns("$A").ColumnWidth = 25     '
        
        .Columns("$B").ColumnWidth = 50
        .Columns("$C").ColumnWidth = 50
        .Columns("$D").ColumnWidth = 30
        .Columns("$E").ColumnWidth = 10
        
        .Columns("$F").ColumnWidth = 15
        
        .Columns("$G").ColumnWidth = 25
        
        .Columns("$H").ColumnWidth = 50
        .Columns("$I").ColumnWidth = 50
        .Columns("$J").ColumnWidth = 30
        .Columns("$K").ColumnWidth = 10
        
        .Columns("$L").ColumnWidth = 100
    End With
XT: Exit Sub

EH: ReRaiseError Err, mModuleName & "FormatTargetWorksheet"
    Resume      ' for debugging only
End Sub

Private Sub DeleteOldAnalysisWorksheets(ByVal DeleteFromThis As Boolean)
    If DeleteFromThis Then
        DeleteTargetWorksheet TargetWorksheetNameThis
        DeleteTargetWorksheet TargetWorksheetNameFilesFromThis
    End If

    DeleteTargetWorksheet TargetWorksheetNameList
    DeleteTargetWorksheet TargetWorksheetNameFilesFromList
    
    DeleteTargetWorksheet TargetWorksheetNameErrors
End Sub
