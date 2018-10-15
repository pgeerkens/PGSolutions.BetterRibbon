Public Enum Token
    ScanError
    EOT
    Identifier
    StringLiteral
    Number
    BinOp
    Unop
    Equals
    Comma
    SemiColon
    Bang
    OpenParen
    CloseParen
    ExternRef
    OpenBrace
    CloseBrace
End Enum

Public Class LinksAnlyzer
    Private Const mModuleName As String = "LinksAnalyzer."
    Private Const TargetName As String = "Link Analysis"

    ''' <summary>TODO</summary>
    Public Sub ListExternalLinksActiveWkbk(
    ByVal ShowErrors As Boolean,
    ByVal IncludeHyperlinks As Boolean
)
        Const MethodName As String = mModuleName & "ListExternalLinksActiveWkbk"

        On Error GoTo EH
        If ActiveWorkbook Is Nothing Then Exit Sub

        Application.ScreenUpdating = False
        Dim ProtectStructure As Boolean : ProtectStructure = ActiveWorkbook.ProtectStructure
        If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Protect , False

    DeleteOldAnalysisWorksheets True

    With New ExternalLinks
            .ExtendFromWorkBook ActiveWorkbook, TargetWorksheetNameThis

        If .Count = 0 Then
                MsgBox LanguageStrings.GetString("NoLinksFound"),
                vbOKOnly Or vbInformation, MethodName
        Else
                WriteFileList.ExternalFiles, TargetWorksheetNameFilesFromThis
            WriteLinksAnalysis.This, TargetWorksheetNameThis, IncludeHyperlinks
            If ShowErrors And .Errors.Count > 0 Then WriteErrors TargetWorksheetNameErrors, .Errors

            On Error Resume Next
                ActiveWorkbook.Sheets(TargetWorksheetNameThis).Select

                On Error GoTo EH
            End If
        End With

XT:     ActiveWorkbook.Protect , ProtectStructure
    Application.ScreenUpdating = True
        Exit Sub

EH:     ErrorUtils.ReRaiseError Err, MethodName
    Resume XT
        Resume      ' for debugging only
    End Sub

    ''' <summary>TODO</summary>
    Public Sub ListExternalLinksWkbkList(
    ByVal rng As Range,
    ByVal ShowErrors As Boolean,
    ByVal IncludeHyperlinks As Boolean
)
        Const MethodName As String = mModuleName & "ListExternalLinksWkbkList"

        On Error GoTo EH
        If ActiveWorkbook Is Nothing Then Exit Sub

        Application.ScreenUpdating = False
        Dim ProtectStructure As Boolean : ProtectStructure = ActiveWorkbook.ProtectStructure
        If ActiveWorkbook.ProtectStructure Then ActiveWorkbook.Protect , False
    Application.Calculation = xlCalculationManual

        Dim row As Long
        Dim Files As Collection :   Set Files = New Collection
    For row = 1 To rng.Rows.Count
            Dim sCell As String : sCell = rng.Cells(row, 1)

            On Error GoTo ErrHandlerDuplicateFile
            If sCell <> "" Then Files.Add sCell, sCell

        On Error GoTo EH
            DoEvents
        Next row

        DeleteOldAnalysisWorksheets False

    If Files.Count = 0 Then
            MsgBox LanguageStrings.GetString("NoFilesFound"),
            vbOKOnly Or vbInformation, MethodName
    Else
            WriteFilesLinks Files, ShowErrors, IncludeHyperlinks
    End If

XT:     Application.Calculation = xlCalculationAutomatic
        Application.Calculate
        ActiveWorkbook.Protect , ProtectStructure
    Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub

ErrHandlerDuplicateFile:
        If Err.Number = 457 Then Resume Next 'This key is already associated with an element of this collection

EH:     ErrorUtils.ReRaiseError Err, MethodName
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
    Private Sub WriteFilesLinks(
    ByVal Files As Collection, ByVal ShowErrors As Boolean,
    ByVal IncludeHyperlinks As Boolean
)
        Const MethodName As String = mModuleName & "WriteFilesLinks"

        On Error GoTo EH
        Dim Links As ExternalLinks :   Set Links = New ExternalLinks
    With Links
            Dim FileName As Variant
            For Each FileName In Files
                Application.StatusBar = "Opening " & FileName

                Application.DisplayAlerts = False
                Dim wb As Workbook :   Set wb = Application.Workbooks.Open(FileName, False, True)
            
            Application.DisplayAlerts = True
                If Not wb Is Nothing Then .ExtendFromWorkBook wb

            wb.Close False
            Set wb = Nothing
            DoEvents
NextFile:
            Next FileName

WrapUp:
            If .Count = 0 And .Errors.Count = 0 Then
                MsgBox LanguageStrings.GetString("NoLinksFound"),
                vbOKOnly Or vbInformation, MethodName
        Else
                WriteFileList.ExternalFiles, TargetWorksheetNameFilesFromList
            WriteLinksAnalysis.This, TargetWorksheetNameList, IncludeHyperlinks
            If ShowErrors And .Errors.Count > 0 Then WriteErrors TargetWorksheetNameErrors, .Errors

            On Error Resume Next
                ActiveWorkbook.Sheets(TargetWorksheetNameList).Select

                On Error GoTo EH
            End If
        End With

XT:     Exit Sub

EH:     If Not wb Is Nothing Then wb.Close False
    If Err.Number = 1004 Then
            Dim mbr As VbMsgBoxResult : mbr = ErrorUtils.MsgBoxAbortRetryIgnore(Err, "WorkBook Not FOund")
            Select Case mbr
                Case VbMsgBoxResult.vbAbort : Links.Errors.AddFileAccessError CStr(FileName), "Abort"
                                        Resume WrapUp
                Case VbMsgBoxResult.vbRetry : Resume
                Case VbMsgBoxResult.vbIgnore : Links.Errors.AddFileAccessError CStr(FileName), "Ignore"
                                        Resume NextFile
            End Select
        End If
        ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    ''' <summary>TODO</summary>
    Public Sub WriteErrors(ByVal TargetName As String, ByVal List As ParseErrors)
        Const MethodName As String = mModuleName & "WriteErrors"
        Const Message As String = "Writing errors to worksheet ... "

        On Error GoTo EH
        Dim TargetWS As Worksheet :   Set TargetWS = CreateTargetWorksheet(ActiveWorkbook, TargetName)
    Dim LastRow As Long : LastRow = 2

        Dim pe As Variant, i As Long : i = 0
        For Each pe In List
            i = i + 1
            LastRow = LastRow + 1
            With pe
                Dim ws As Worksheet :   Set ws = TargetWS
            Dim col As Long : col = 0
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .CellRef.Cell
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .CellRef.TabName
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .CellRef.FileName
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .Condition
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .CharPosition
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = "'" & .Formula
                col = col + 1 : TargetWS.Cells(LastRow, col).Value2 = .CellRef.Path
            End With

            Dim Completion As Long : Completion = i * 100 / List.Count
            Application.StatusBar = Message & "(" & CStr(Completion) & "%)"
            DoEvents
        Next pe
        With LinksErrorColumns : InitializeTargetWorksheet TargetWS, LastRow, .This: End With

XT:     Exit Sub

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    Private Property Get LinksErrorColumns() As ColumnHeaders
    With New ColumnHeaders
    .Initialize _
            "Source" & vbNewLine & "Cell",
            "Source" & vbNewLine & "Worksheet",
            "Source" & vbNewLine & "FileName",
            "Error" & vbNewLine & "Description",
            "Char Pos",
            "Source" & vbNewLine & "Formula",
            "External" & vbNewLine & "Path"
        Set LinksErrorColumns = .This
    End With
    End Property

    Private Property Get LinksAnalysisColumns() As ColumnHeaders
    With New ColumnHeaders
    .Initialize _
            "Links Target",
            "External" & vbNewLine & "Path",
            "External" & vbNewLine & "FileName",
            "External" & vbNewLine & "Worksheet",
            "External" & vbNewLine & "Cell",
            "Link" & vbNewLine & "Type",
            "Links Source",
            "Source" & vbNewLine & "Path",
            "Source" & vbNewLine & "FileName",
            "Source" & vbNewLine & "Worksheet",
            "Source" & vbNewLine & "Cell",
            "Source Formula"
        Set LinksAnalysisColumns = .This
    End With
    End Property

    ''' <summary>TODO</summary>
    ''' <returns>Returns the number of the LastRow added to the target worksheet.</returns>
    Private Sub WriteLinksAnalysis(
    ByVal List As ExternalLinks,
    ByVal TargetName As String,
    Optional ByVal IncludeHyperlinks As Boolean = False
)
        Const MethodName As String = mModuleName & "WriteLinksAnalysis"
        Const HyperLinksMessage As String = "Writing Hyperlinks: "

        On Error GoTo EH
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False

        If List.Count > 0 Then
            Dim d As ITwoDimensionalLookup :   Set d = List
        Dim MaxCol As Long : MaxCol = d.ColsCount

            Dim ws As Worksheet :   Set ws = ClearTargetWorksheet(ActiveWorkbook, TargetName)
        ws.Rows(MaxCol).EntireRow.NumberFormat = "Text"

            Dim FirstCell As Range :       Set FirstCell = ws.Cells(3, 1)
        Dim LastCell As Range :       Set LastCell = ws.Cells(d.RowsCount + 2, MaxCol)
        Dim SheetData As Range :       Set SheetData = ws.Range(FirstCell, LastCell)
        
        API.CopyToRange d, SheetData

        If IncludeHyperlinks Then
                Dim i As Long, FileName As String, Range As String
                For i = 1 To List.Count
                    With List.ItemByIndex(i)
                        WriteLinkHyperlink SheetData.Cells(i, 1), .Path, .FileName, .TabName, .Cell
                    If .LinkType = "Cell Reference" Then _
                        WriteLinkHyperlink SheetData.Cells(i, 7), .SourcePath & "\", .SourceFile, .SourceTab, .SourceCell
                End With

                    If SheetData.Hyperlinks.Count > 65500 Then Exit For
                    Dim Completion As Long : Completion = i * 100 / List.Count
                    Application.StatusBar = HyperLinksMessage & "(" & CStr(Completion) & "%)"
                    DoEvents
                Next i
            End If

            With LinksAnalysisColumns : InitializeTargetWorksheet ws, LastCell.row, .This: End With
        End If

XT:     Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub

EH:     Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    Private Sub WriteLinkHyperlink(ByVal Anchor As Range, ByVal Path As String, ByVal FileName As String, ByVal TabName As String, CellAddress As String)
        Const MethodName As String = mModuleName & "WriteLinkHyperlink"
        Const HyperlinkInfo As String = vbNewLine & vbNewLine & "Ctrl-LeftClick to select this cell;" &
                                        vbNewLine & vbNewLine & "Left Click to follow link"
        On Error GoTo EH
        Dim Address As String : Address = Path & FileName
        Dim SubAddress As String
        If CellAddress <> "" Then
            SubAddress = "'" & TabName & "'!" & IIf(CellAddress = "#REF!", "", CellAddress)
        Else
            SubAddress = ""
        End If
        If CellAddress <> "#REF!" And CellAddress <> "" Then SubAddress = Application.ConvertFormula(SubAddress, xlA1, xlA1, xlAbsolute)

        With Anchor.Worksheet.Hyperlinks
            On Error Resume Next    ' Limit of 65535 links or so - blast through anyways
            If .Count < 65535 Then .Add Anchor, "file://" & Address, SubAddress,
                "file://" & Address & SubAddress & HyperlinkInfo,
                Address & SubAddress
    End With
XT:     Exit Sub

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    Private Sub WriteFileList(ByVal Files As ExternalFiles, ByVal Target As String)
        Const MethodName As String = mModuleName & "WriteFileList"
        Const Message As String = "Writing links to worksheet ... "

        On Error GoTo EH
        With ClearTargetWorksheet(ActiveWorkbook, Target)
            Dim FileName As Variant, row As Long : row = 0
            For Each FileName In Files
                row = row + 1 : .Cells(row, 1).Value = FileName
                WriteLinkHyperlink.Cells(row, 1), "", CStr(FileName), "", ""
            Dim Completion As Long : Completion = row * 100 / Files.Count
                Application.StatusBar = "Writing File List: " & "(" & CStr(Completion) & "%)"
            Next FileName
            .Range("$A:$A").EntireColumn.AutoFit
        End With

XT:     Exit Sub

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    ''' <summary>TODO</summary>
    Private Sub DeleteTargetWorksheet(ByVal Name As String)
        Const MethodName As String = mModuleName & "DeleteTargetWorksheet"

        On Error GoTo EH
        Application.DisplayAlerts = False
        Sheets(Name).Delete
        Application.DisplayAlerts = True
XT:     Exit Sub

EH:     If Err.Number = 9 Then Resume Next
        ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    ''' <summary>TODO</summary>
    Private Function ClearTargetWorksheet(
    ByVal wb As Workbook,
    ByVal Name As String
) As Excel.Worksheet
        Const MethodName As String = mModuleName & "ClearTargetWorksheet"

        On Error GoTo EH
    Set ClearTargetWorksheet = wb.Worksheets(Name)
    
    If Not ClearTargetWorksheet Is Nothing Then DeleteTargetWorksheet Name
    Set ClearTargetWorksheet = CreateTargetWorksheet(wb, Name)

XT:     Exit Function

EH:     If Err.Number = 9 Then Resume Next
        ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    ''' <summary>TODO</summary>
    Private Function CreateTargetWorksheet(ByVal wb As Workbook, ByVal Name As String) As Excel.Worksheet
        Const MethodName As String = mModuleName & "CreateTargetWorksheet"

        On Error GoTo EH
        DeleteTargetWorksheet Name

    Dim TargetWS As Worksheet :   Set TargetWS = wb.Worksheets.Add(before:=wb.Worksheets(1))
    TargetWS.Name = Name
    
    Set CreateTargetWorksheet = TargetWS
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    ''' <summary>TODO</summary>
    Private Sub InitializeTargetWorksheet(ByVal TargetWS As Excel.Worksheet, ByVal LastRow As Long, ByVal ColHeaders As ColumnHeaders)
        Const MethodName As String = mModuleName & "InitializeTargetWorksheet"

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

XT:     Exit Sub

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub
    ''' <summary>TODO</summary>
    Private Sub FormatTargetWorksheet(ByVal TargetWS As Worksheet)
        Const MethodName As String = mModuleName & "FormatTargetWorksheet"

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
XT:     Exit Sub

EH:     ReRaiseError Err, MethodName
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

End Class
