Option Explicit On
Option Compare Binary

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace PGSolutions.ExcelRibbon2013

    Public Class ExportVba
        ''' <summary>Extracts VBA modules from current EXCEL workbook to sibling directory 'src'..</summary>
        Public Sub ExportModulesCurrentProject(Optional ByVal DestIsSrc As Boolean = True)
            Try
                Application.Cursor = xlWait
                Application.ScreenUpdating = False
                Call (New ProjectFilterExcel).ExtractOpenProject(ActiveWorkbook, DestIsSrc)
            Finally
                Application.StatusBar = False
                Application.ScreenUpdating = True
                Application.Cursor = xlDefault
            End Try
        End Sub

        ''' <summary>Extracts VBA modules from a selected EXCEL workbook to the sibling directory 'src'.</summary>
        Public Sub ExportModules(Optional ByVal DestIsSrc As Boolean = True)
            Try
                Dim list As ProjectFilters = New ProjectFilters
                With Application.FileDialog(msoFileDialogFilePicker)
                    .AllowMultiSelect = False
                    .ButtonName = "Export"
                    .Title = "Select VBA Project(s) to Export From"

                    .Filters.Clear
                    Dim i As Long
                    For i = 1 To list.Count
                        Dim Item As IProjectFilter = list.ItemByIndex(i)
                        .Filters.Add(Item.Description, Item.Extensions)
                        DoEvents
                    Next i

                    If .Show <> 0 Then
                        Application.Cursor = xlWait
                        Application.ScreenUpdating = False

                        On Error GoTo EH_Busy
                        Dim Filter As IProjectFilter = list.ItemByIndex(.FilterIndex)
                        Filter.ExtractProjects(.SelectedItems, DestIsSrc)
                    End If
                End With
            Catch ex As Exception
            Finally
                Application.StatusBar = False
                Application.ScreenUpdating = True
                Application.Cursor = xlDefault
            End Try
        End Sub
    End Class
End Namespace
