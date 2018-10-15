Public Class ParseError

    Private mCellRef As InternalCellRef
    Private mFormula As String
    Private mCondition As String
    Private mCharPosition As Long

    Friend Function Initialize(ByVal CellRef As InternalCellRef, ByVal Formula As String,
    ByVal CharPosition As Long,
    ByVal Condition As String
) As ParseError
    Set mCellRef = CellRef
    mFormula = Formula
        mCondition = Condition
        mCharPosition = CharPosition
    Set Initialize = Me
End Function

    Friend Property Get This() As ParseError : Set This = Me:                  End Property

    Public Property Get CellRef() As InternalCellRef : Set CellRef = mCellRef:         End Property

    Public Property Get CharPosition() As Long :          CharPosition = mCharPosition:   End Property

    Public Property Get Condition() As String :           Condition = mCondition:         End Property

    Public Property Get Formula() As String :             Formula = mFormula:             End Property

End Class
