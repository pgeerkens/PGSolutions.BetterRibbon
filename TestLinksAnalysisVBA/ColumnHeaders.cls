VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCollection         As Collection

Public Property Get Count() As Long:                        Count = mCollection.Count:              End Property
Public Function ItemByIndex(ByVal Index As Long) As String: ItemByIndex = mCollection.Item(Index):  End Function
Public Function ItemById(ByVal ID As String) As String:     ItemById = mCollection.Item(ID):        End Function

Private Sub Class_Initialize(): Set mCollection = New Collection:                                   End Sub
Private Sub Class_Terminate():  ClearCollection mCollection: Set mCollection = Nothing:   End Sub

Public Property Get This() As ColumnHeaders:                           Set This = Me:               End Property

Private Sub Add(ByVal ColumnCaption As String)
    mCollection.Add ColumnCaption, ColumnCaption
End Sub

Public Sub Initialize(ParamArray ColumnNames() As Variant)
    On Error GoTo EH
    Dim ColumnName As Variant
    For Each ColumnName In ColumnNames:  Add ColumnName: Next ColumnName
XT: Exit Sub
EH: ReRaiseError Err, TypeName(Me) & ".Initialize"
End Sub
