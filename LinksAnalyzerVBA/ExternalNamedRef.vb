Public Class ExternalNamedRef
    Private Const mModuleName As String = "ExternalCellRef."

    Implements ICellRef

    Private mIsInitialized As Boolean
    Private mPath As String
    Private mFile As String
    Private mTab As String
    Private mCell As String
    Private mText As String

    Private mFormula As String
    Private mSourcePath As String
    Private mSourceFile As String
    Private mSourceTab As String
    Private mSourceCell As String

    ''' <summary>TODO</summary>
    Public Property Get ICellRef_Path() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "Path"
    ICellRef_Path = mPath
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_FileName() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "FileName"
    ICellRef_FileName = mFile
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_TabName() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "TabName"
    ICellRef_TabName = mTab
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_Cell() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "Cell"
    ICellRef_Cell = mCell
End Property

    ''' <summary>TODO</summary>
    Public Property Get ICellRef_SourcePath() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "SourcePath"
    ICellRef_SourcePath = mSourcePath
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_SourceFile() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "SourceFile"
    ICellRef_SourceFile = mSourceFile
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_SourceTab() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "SourceTab"
    ICellRef_SourceTab = mSourceTab
End Property
    ''' <summary>TODO</summary>
    Public Property Get ICellRef_SourceCell() As String
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "SourceCell"
    ICellRef_SourceCell = mSourceCell
End Property

    Public Property Get ICellRef_Formula() As String
    ICellRef_Formula = mFormula
End Property
    Public Property Let Formula(Value As String)
    mFormula = Value
End Property

    Public Property Get ICellRef_IsNamedRangeRef() As Boolean
    ICellRef_IsNamedRangeRef = True
End Property

    Public Property Get ICellRef_LinkType() As String
    ICellRef_LinkType = "Named Range"
End Property

    ''' <summary>TODO</summary>
    Public Property Get This() As ExternalNamedRef
        Set This = Me
End Property

    ''' <summary>TODO</summary>
    Public Function Parse(ByVal Path As String, ByVal TextIn As String,
    ByVal CellRef As InternalCellRef, ByVal Formula As String
) As Boolean
        Const MethodName As String = mModuleName & "Parse"

        On Error GoTo EH
        mText = Path & "!" & TextIn
        If VBA.mID$(mText, 1, 1) = "'" Then
            Dim indexBra As Long : indexBra = VBA.InStr(1, mText, "[") : If indexBra = 0 Then GoTo XT
            Dim indexKet As Long : indexKet = VBA.InStr(indexBra, mText, "]") : If indexKet = 0 Then GoTo XT
            Dim indexBang As Long : indexBang = VBA.InStr(indexKet, mText, "'!") : If indexBang = 0 Then GoTo XT

            mPath = VBA.mID$(mText, 2, indexBra - 2)
            mFile = VBA.mID$(mText, indexBra + 1, indexKet - indexBra - 1)
            mTab = VBA.mID$(mText, indexKet + 1, indexBang - indexKet - 1)
            mCell = VBA.mID$(mText, indexBang + 2, VBA.Len(mText) - indexBang - 1)
            '    mCell = VBA.Replace(mCell, "$", "")

            mSourcePath = CellRef.Path & "\"
            mSourceFile = CellRef.FileName ' SourceFile
            mSourceTab = CellRef.TabName ' Source.Parent
            mSourceCell = CellRef.Cell '.Name
            mFormula = Formula
        End If
        mIsInitialized = True
        Parse = mIsInitialized
XT:
        Exit Function
EH:     ReRaiseError Err, MethodName
End Function

    Private Sub Class_Initialize()
        mIsInitialized = False
    End Sub

End Class
