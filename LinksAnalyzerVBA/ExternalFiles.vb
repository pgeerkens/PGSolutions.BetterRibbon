Public Class ExternalFiles

    Private Const mModuleName As String = "ExternalFiles."

    Private mCollection As Collection

    Public Property Get Count() As Long :                         Count = mCollection.Count:              End Property
    Public Function ItemByIndex(ByVal Index As Long) As String : ItemByIndex = mCollection.Item(Index) : End Function
    Public Function ItemById(ByVal ID As String) As String : ItemById = mCollection.Item(ID) : End Function

    Private Sub Class_Initialize() :  Set mCollection = New Collection:                                   End Sub
    Private Sub Class_Terminate() : Utilities.ClearCollection mCollection:  Set mCollection = Nothing:   End Sub

    ''' <summary>For Each enumerator.</summary>
    ''' <remarks>
    ''' If compilation fails, restore Attribute lines by exporting to file; uncommenting in text editor; and re-importing to project
    ''' </remarks>
    Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'    Attribute NewEnum.VB_UserMemId = -4
'    Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = mCollection.[_NewEnum]
End Property

    Friend Function Add(ByVal FileName As String) As ExternalFiles
        Const MethodName As String = mModuleName & "Add"
        On Error GoTo EH
        mCollection.Add FileName, FileName
XT:     Exit Function
EH:     If Err.Number = ErrorUtils.ErrorNumAlreadyInCollection Then Resume XT
        ReRaiseError Err, MethodName
End Function
End Class
