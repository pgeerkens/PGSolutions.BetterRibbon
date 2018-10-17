'////////////////////////////////////////////////////////////////////////////////////////////////////
'//                                Copyright (c) 2017-8 Pieter Geerkens                              //
'////////////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit On
Option Compare Binary

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports Microsoft.Office.Tools.Excel
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel

Namespace PGSolutions.ExcelRibbon2013
    Public Class ProjectFilters
        ''' <summary>Creates a new instance of the collection.</summary>
        Public Sub ProjectFilters()
            _collection = New Collection

            _collection.Add((New ProjectFilterExcel).IProjectFilter_Initialize(Key, "*.xlsm;*.xlsb;*.xls"), "MS-Excel Workbooks")
            _collection.Add((New ProjectFilterExcel).IProjectFilter_Initialize(Key, "*.xla"), "MS-Excel Add-Ins")

            '    If (New AccessWrapper).IsAccessSupported Then
            '    _collection.Add((New ProjectFilterAccess).IProjectFilter_Initialize(Key, "*.mdb;*.accdb"), "MS-Access Databases")
            '    _collection.Add((New ProjectFilterAccess).IProjectFilter_Initialize(Key, "*.mda;*.accda"), "MS-Access Add-Ins")
            '    End If
        End Sub

        ''' <summary>Returns the number of items in the collection.</summary>
        Public Property Count() As Long
            Get
                Return _collection.Count
            End Get
        End Property

        ''' <summary>Returnes the item at the specified Index of the collection.</summary>
        Public Function ItemByIndex(ByVal Index As Long) As IProjectFilter
            ItemByIndex = _collection.Item(Index)
        End Function
        ''' <summary>Returns the item with the specified Key in the collection.</summary>
        Public Function ItemById(ByVal ID As String) As IProjectFilter
            ItemById = _collection.Item(ID)
        End Function

        Private _collection As Collection
    End Class

    Partial Public Class ProjectFilterExcel

        ''' <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        Private Sub ExtractClosedProject(
            ByVal app As Excel.Application,
            ByVal FileName As String,
            ByVal DestIsSrc As Boolean
        )
            Dim WkBk As Excel.Workbook = app.Workbooks.Open(FileName,
                    UpdateLinks:=False, ReadOnly:=True, AddToMru:=False, Editable:=False)

            Try
                ExtractOpenProject(WkBk, DestIsSrc)
            Finally
                If Not WkBk Is Nothing Then WkBk.Close
            End Try
        End Sub

        ''' <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        Public Sub ExtractOpenProject(ByVal WkBk As Workbook, ByVal DestIsSrc As Boolean)
            CodeExport.ExtractModulesByProject(WkBk.VBProject, CreateDirectory(WkBk.FullName, DestIsSrc))
        End Sub
    End Class
End Namespace