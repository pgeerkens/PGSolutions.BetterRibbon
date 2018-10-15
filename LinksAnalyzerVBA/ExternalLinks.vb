Public Class ExternalLinks
    Private Const mModuleName As String = "ExternalLinks."

    Private mIsInitialized As Boolean
    Private mCollection As Collection
    Private mParseErrors As ParseErrors
    Private mLocation As InternalCellRef
    Private mLexer As LinksLexer

    Private mFiles As ExternalFiles

    Implements ITwoDimensionalLookup

    ''' <summary>TODO</summary>
    Public Property Get Count() As Long
    Count = mCollection.Count
End Property

    ''' <summary>TODO</summary>
    Public Property Get ItemByKey(ByVal Key As String) As ICellRef
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "ItemByKey"
    Set ItemByKey = mCollection.Item(Key)
End Property
    ''' <summary>TODO</summary>
    Public Property Get ItemByIndex(ByVal Index As Long) As ICellRef
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "ItemByIndex"
    Set ItemByIndex = mCollection.Item(Index)
End Property

    ''' <summary>For Each enumerator.</summary>
    ''' <remarks>
    ''' If compilation fails, restore Attribute lines by exporting to file; uncommenting in text editor; and re-importing to project
    ''' </remarks>
    Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "240"
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "240"
    
    Set NewEnum = mCollection.[_NewEnum]
End Property

    ''' <summary>TODO</summary>
    Public Property Get This() As ExternalLinks
        Set This = Me
End Property

    Public Property Get Errors() As ParseErrors
        Set Errors = mParseErrors
End Property

    ''' <summary>TODO</summary>
    Public Function Parse(
    ByVal Lexer As LinksLexer,
    ByVal CellRef As InternalCellRef
) As ExternalLinks
        Const MethodName As String = mModuleName & "Initialize"

        On Error GoTo EH
    Set mLocation = CellRef
    
    With Lexer
            Do
                Dim TokenText As String
                Dim Token As Token : Token = .Scan(TokenText)
                If Token = EOT Then Exit Do

                Select Case Token
              '  Case Else:
                    ' NO-OP
                    Case ExternRef
                        ParseExternRef.This, TokenText, CellRef
                Case ScanError
                        mParseErrors.Add.RaiseError(CellRef, "Unknown token found")
                        GoTo XT
                End Select
            Loop
        End With

        mIsInitialized = True
    Set Parse = Me
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    Public Property Get ITwoDimensionalLookup_RowsCount() As Long :  ITwoDimensionalLookup_RowsCount = Count: End Property
    Public Property Get ITwoDimensionalLookup_ColsCount() As Long :  ITwoDimensionalLookup_ColsCount = 12:    End Property

    ''' <summary>TODO</summary>
    Private Function ITwoDimensionalLookup_Item(ByVal RowNo As Long, ByVal ColNo As Long) As Variant
        Const MethodName As String = mModuleName & "AsArray"

        On Error GoTo EH
        With ItemByIndex(RowNo + 1)
            Select Case ColNo
                Case 0 : ITwoDimensionalLookup_Item = .Path & .FileName

                Case 1 : ITwoDimensionalLookup_Item = .Path
                Case 2 : ITwoDimensionalLookup_Item = .FileName
                Case 3 : ITwoDimensionalLookup_Item = .TabName
                Case 4 : ITwoDimensionalLookup_Item = .Cell

                Case 5 : ITwoDimensionalLookup_Item = .LinkType

                Case 6 : ITwoDimensionalLookup_Item = .SourcePath & .SourceFile

                Case 7 : ITwoDimensionalLookup_Item = .SourcePath
                Case 8 : ITwoDimensionalLookup_Item = .SourceFile
                Case 9 : ITwoDimensionalLookup_Item = .SourceTab
                Case 10 : ITwoDimensionalLookup_Item = .SourceCell

                Case 11 : ITwoDimensionalLookup_Item = "'" & .Formula
            End Select
        End With

XT:     Exit Function

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    ''' <summary>TODO</summary>
    Public Function AsArray(Optional ByVal ShowCompletion As Boolean = False) As Variant()
        Const MethodName As String = mModuleName & "AsArray"
        Const Message As String = "Writing links to worksheet ... "

        On Error GoTo EH
        Dim i As Long, col As Long, Data() As Variant
        If Count > 0 Then
            ReDim Data(0 To Count - 1, 0 To 11) As Variant
        For i = 0 To Count - 1
                With ItemByIndex(i + 1)
                    col = 0
                    Data(i, col) = .Path & .FileName : col = col + 1

                    Data(i, col) = .Path : col = col + 1
                    Data(i, col) = .FileName : col = col + 1
                    Data(i, col) = .TabName : col = col + 1
                    Data(i, col) = .Cell : col = col + 1

                    Data(i, col) = .LinkType : col = col + 1

                    Data(i, col) = .SourcePath & .SourceFile : col = col + 1

                    Data(i, col) = .SourcePath : col = col + 1
                    Data(i, col) = .SourceFile : col = col + 1
                    Data(i, col) = .SourceTab : col = col + 1
                    Data(i, col) = .SourceCell : col = col + 1

                    Data(i, col) = "'" & .Formula : col = col + 1
                End With

                Dim Completion As Long : Completion = i * 100 / Count
                If ShowCompletion Then Application.StatusBar = Message & "(" & CStr(Completion) & "%)"
                DoEvents
            Next i
        End If
        AsArray = Data

XT:     Exit Function

EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    ''' <summary>TODO</summary>
    Friend Sub ExtendFromWorkBook(ByVal wb As Workbook,
    Optional ByVal ExcludedWorksheetName As String = ""
)
        Const MethodName As String = mModuleName & "ExtendFromWorkSheet"

        On Error GoTo EH
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If ws.Name <> ExcludedWorksheetName Then _
                ExtendFromWorkSheet ws, "Searching " & wb.Name & "[" & ws.Name & "] ... (??%)"
        DoEvents
        Next ws

        Dim Source As Excel.Name
        Dim cFormula As String
        For Each Source In wb.Names
            cFormula = Source.RefersTo

            If Len(cFormula) > 0 Then
                mLexer.LoadText cFormula
            If VBA.Left$(cFormula, 1) = "=" Then Parse mLexer.This, NewWorkbookNameRef(wb, Source)
        End If
            DoEvents
NextName:
        Next Source

XT:     Exit Sub
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    ''' <summary>TODO</summary>
    Friend Sub ExtendFromWorkSheet(ByVal ws As Worksheet, ByVal MessageText As String)
        Const MethodName As String = mModuleName & "ExtendFromWorkSheet"

        On Error GoTo EH

        Dim UsedRange As Range :  Set UsedRange = ws.UsedRange
    Dim cl As Range, cFormula As String
        Dim ColNo As Long
        For ColNo = 1 To UsedRange.Columns.Count
            Dim Percentage As Long : Percentage = 100 * ColNo / UsedRange.Columns.Count
            Application.StatusBar = Replace(MessageText, "??", CStr(Percentage))

            Dim LastRowNo As Long : LastRowNo = ws.Cells(ws.Rows.Count, ColNo).End(xlUp).row
            Dim RowNo As Long
            For RowNo = 1 To LastRowNo
            Set cl = ws.Cells(RowNo, ColNo)
            cFormula = cl.Formula
                If Len(cFormula) > 0 Then
                    mLexer.LoadText cFormula
                If VBA.Left$(cFormula, 1) = "=" Then Parse mLexer.This, NewCellRef(ws, cl)
            End If
                DoEvents
            Next RowNo
        Next ColNo

XT:     Exit Sub
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Sub

    Friend Property Get ExternalFiles() As ExternalFiles
        Set ExternalFiles = mFiles
End Property

    Private Function NewCellRef(ByVal ws As Worksheet, ByVal cl As Range) As InternalCellRef
        Const MethodName As String = mModuleName & "NewCellRef"

        On Error GoTo EH
        With New InternalCellRef
            .Initialize ws.Parent.Path, ws.Parent.Name, ws.Name, cl.Address
        Set NewCellRef = .This
    End With

XT:     Exit Function
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    Private Function NewWorkbookNameRef(ByVal wb As Workbook, ByVal NamedRange As Name) As InternalCellRef
        Const MethodName As String = mModuleName & "NewCellRef"

        On Error GoTo EH
        With New InternalCellRef
            Dim SheetName As String
            If NamedRange.Parent Is wb Then
                SheetName = "<workbook>"
            Else
                SheetName = NamedRange.Parent.Name
            End If
            .Initialize wb.Path, wb.Name, SheetName,
                Replace(Replace(NamedRange.Name, "'" & SheetName & "'!", ""), SheetName & "!", "") _
                , True
        Set NewWorkbookNameRef = .This
    End With

XT:     Exit Function
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function


    ''' <summary>TODO</summary>
    Private Function ParseExternRef(
    ByVal Lexer As LinksLexer,
    ByVal Path As String,
    ByVal CellRef As InternalCellRef
) As Token
        Const MethodName As String = mModuleName & "ParseExternRef"

        On Error GoTo EH
        Dim TokenText As String
        With Lexer
            Dim t As Token : t = .Scan(TokenText)
            If t = Bang Then t = .Scan(TokenText)
            If t <> Identifier Then
                ParseExternRef = mParseErrors.Add(.RaiseError(mLocation, "Expected Identifier"))
                GoTo XT
            End If
        End With

        If CellRef.IsNamedRangeRef Then
            With New ExternalNamedRef
                If .Parse(Path, TokenText, CellRef, Lexer.Text) Then
                    Add.This
                    mFiles.Add.This.ICellRef_Path & .This.ICellRef_FileName
            End If
            End With
        Else
            With New ExternalCellRef
                If .Parse(Path, TokenText, CellRef, Lexer.Text) Then
                    Add.This
                    mFiles.Add.This.ICellRef_Path & .This.ICellRef_FileName
            End If
            End With
        End If

XT:
        Exit Function
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    ''' <summary>TODO</summary>
    Private Function Add(ByVal Item As ICellRef)
        mCollection.Add Item, CStr(mCollection.Count + 1)
End Function

    Private Sub Class_Initialize()
        mIsInitialized = False
    Set mCollection = New Collection
    Set mFiles = New ExternalFiles
    Set mLexer = New LinksLexer
    Set mParseErrors = New ParseErrors
End Sub

End Class
