////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms; //  *** TODO *** THis needs to be moved into ExcelRibbon.

using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

using PGSolutions.LinksAnalyzer.Interfaces;
using System.Runtime.InteropServices;

namespace PGSolutions.LinksAnalyzer {
    [CLSCompliant(false)]
    public static class ExcelLinksExtensions {
        public static void WriteLinks(this Workbook wb) {
            wb.DeleteTargetWorksheet(LinksSheetName);
            wb.DeleteTargetWorksheet(FilesSheetName);
            wb.DeleteTargetWorksheet(ErrorsSheetName);

            wb.WriteLinks(new ExternalLinks(wb, ""));
        }

        public static void WriteLinks(this Workbook wb, VBA.Collection nameList) {
            wb.DeleteTargetWorksheet(LinksSheetName);
            wb.DeleteTargetWorksheet(FilesSheetName);
            wb.DeleteTargetWorksheet(ErrorsSheetName);

            wb.WriteLinks(new ExternalLinks(wb.Application, nameList));
        }

        public const string LinksSheetName  = "Links Analysis";
        public const string FilesSheetName  = "Linked Files";
        public const string ErrorsSheetName = "Links Errors";

        internal static void WriteLinks(this Workbook wb, IExternalLinks links) {
            if(links.Count == 0  && links.Errors.Count == 0) {
                MessageBox.Show("No external links found!", "", MessageBoxButtons.OK);
            } else {
                var wsLinks = wb.CreateTargetWorksheet(LinksSheetName);
                var wsFiles = wb.CreateTargetWorksheet(FilesSheetName);
                var wsErrors=wb.CreateTargetWorksheet(ErrorsSheetName);

                wsLinks.WriteLinksAnalysis(links);
                wsFiles.WriteLinksFiles(links.Files);
                wsErrors.WriteLinksErrors(links.Errors);

                wb.Application.StatusBar = false;
            }
        }

        internal static void WriteLinksAnalysis(this Worksheet ws, IExternalLinks links) {
            if(links.Count > 0) {
                var calculation = ws.Application.Calculation;
                try {
                    ws.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                    ws.Application.ScreenUpdating = false;

                    var links2D   = links as ITwoDimensionalLookup;
                    var MaxCol    = links2D.ColsCount;
                    var firstCell = ws.Cells[3,1];
                    var lastCell  = ws.Cells[links2D.RowsCount + 2, MaxCol];
                    var sheetData = ws.Range[firstCell,lastCell] as Excel.Range;

                    ws.Columns[MaxCol].EntireColumn.NumberFormat = "@"; // "Text";  // Formula column
                    links2D.FastCopyToRange(sheetData);

                    ws.InitializeTargetWorksheet(links2D.RowsCount + 2,new List<string>() {
                            "Links Target",     "External\nPath",   "External\nFileName", "External\nWorksheet",
                            "External\nCell",   "Link\nType",       "Source\nType",       "Source\nPath",
                            "Source\nFileName", "Source\nWorksheet","Source\nCell",       "Source Formula"} );
                } finally {
                    ws.Application.ScreenUpdating = true;
                    ws.Application.Calculation = calculation;
                }

            }
        }

        internal static void WriteLinksFiles(this Worksheet ws, IExternalFiles files) {
            var lastRow = 2;
            var i       = 0;
            foreach(var fileName in files.OrderBy(s=>s)) {
                i++; lastRow++;
                ws.Cells[lastRow,1].Value2 = fileName;

                ws.WritePercentageStatus(ws.Name,100*i/files.Count);
                // DoEvents
            }
            ws.InitializeTargetWorksheet(lastRow,new List<string>() {"External FIles"} );
        }

        internal static void WriteLinksErrors(this Worksheet ws, IParseErrors errors) {
            var lastRow = 2;
            var i       = 0;
            foreach(var error in errors) {
                i++; lastRow++;
                var col = 0;
                ws.Cells[lastRow,++col].Value2 = error.CellRef.CellName;
                ws.Cells[lastRow,++col].Value2 = error.CellRef.TabName;
                ws.Cells[lastRow,++col].Value2 = error.CellRef.FileName;
                ws.Cells[lastRow,++col].Value2 = error.Condition;;
                ws.Cells[lastRow,++col].Value2 = error.CharPosition;
                ws.Cells[lastRow,++col].Value2 = $"'{error.Formula}";
                ws.Cells[lastRow,++col].Value2 = error.CellRef.FullPath;

                ws.WritePercentageStatus(ws.Name,100*i/errors.Count);
                // DoEvents
            }

            ws.InitializeTargetWorksheet(lastRow,new List<string>() {
                    "Cell Name", "Tab Name", "File Name", "Error Condition", "Position", "Formula", "Path"} );
        }

        private static void WritePercentageStatus(this Worksheet ws, string sheetName, int percentage) 
            => ws.Application.StatusBar = $"Writing worksheet {sheetName}: ({percentage}%) ... ";

        private static void InitializeTargetWorksheet(this Worksheet ws, long lastRow, IList<string> columnHeaders) {
            var colNo = 0;
            foreach(var columnHeader in columnHeaders) {
                ws.Cells[2,++colNo].Value = columnHeader;
            }

            ws.Range[ws.Cells[2,1], ws.Cells[lastRow,columnHeaders.Count]].Columns.AutoFit();

            ws.Range["A1:D1"].Merge();
            ws.Range["A1"].Formula = $"Link analysis run on {DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}";

            ws.Range["$A$2", ws.Cells[lastRow, columnHeaders.Count]].AutoFilter();
            // *** TODO: *** the Select() call occasionally fails here - reason TBD
            //ws.Application.ActiveWindow.FreezePanes = false;
            //(ws.Rows["$3:$3"].EntireRow as Excel.Range)?.Select();
            //ws.Application.ActiveWindow.FreezePanes = true;
        }

        private static void DeleteTargetWorksheet(this Workbook wb, string sheetName) {
            try {
                wb.Application.DisplayAlerts = false;
                wb.Worksheets[sheetName].Delete();
            } catch ( COMException ex ) {
                /*  NO-OP: Sheet doesn't exist so delete not needed */
            } finally {
                wb.Application.DisplayAlerts = true;
            }
        }

        private static Worksheet CreateTargetWorksheet(this Workbook wb, string sheetName) {
            wb.DeleteTargetWorksheet(sheetName);

            var ws = wb.Worksheets.Add(Before:wb.Worksheets[1]);
            ws.Name = sheetName;
            return ws;
        }
    }
}
