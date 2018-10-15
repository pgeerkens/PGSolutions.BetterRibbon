Public Class ParseErrors
    Private Const mModuleName As String = "ParseErrors."

    Private mCollection As Collection

    Public Property Get Count() As Long :                             Count = mCollection.Count:                  End Property
    Public Function ItemByIndex(ByVal Index As Long) As ParseError :  Set ItemByIndex = mCollection.Item(Index):  End Function
    Public Function ItemById(ByVal ID As String) As ParseError :      Set ItemById = mCollection.Item(ID):        End Function

    Private Sub Class_Initialize() :  Set mCollection = New Collection:                                           End Sub
    Private Sub Class_Terminate() : Utilities.ClearCollection mCollection:  Set mCollection = Nothing:           End Sub

    ''' <summary>For Each enumerator.</summary>
    ''' <remarks>
    ''' If compilation fails, restore Attribute lines by exporting to file; uncommenting in text editor; and re-importing
    ''' </remarks>
    Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
   
    Set NewEnum = mCollection.[_NewEnum]
End Property

    Friend Function Add(ByVal e As ParseError) As Token
        Const MethodName As String = mModuleName & "Add"
        On Error GoTo EH
        mCollection.Add e, CStr(Count + 1)
    Add = ScanError
XT:     Exit Function
EH:     ReRaiseError Err, MethodName
    Resume      ' for debugging only
    End Function

    Friend Sub AddFileAccessError(ByVal FullPath As String, ByVal Action As String)
        Dim CellRef As InternalCellRef
        With New InternalCellRef
        Set CellRef = .Initialize(FullPath, "", "", "")
    End With
        With New ParseError
            .Initialize CellRef, FullPath, 0, "File Not Found"
        Add.This
        End With
    End Sub

End Class
