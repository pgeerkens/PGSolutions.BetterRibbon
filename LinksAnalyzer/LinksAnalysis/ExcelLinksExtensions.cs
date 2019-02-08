﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms; //  *** TODO *** THis needs to be moved into BetterRibbon.

using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    [CLSCompliant(false)]
    public static class ExcelLinksExtensions {
        public static void WriteLinks(this Workbook wb) {
            wb.DeleteTargetWorksheet(LinksSheetName);
            wb.DeleteTargetWorksheet(FilesSheetName);
            wb.DeleteTargetWorksheet(ErrorsSheetName);

            wb.WriteLinks(new ExternalLinks(wb, ""));
        }

        public static void WriteLinks(this Workbook wb, IReadOnlyList<string> nameList) {
            wb.DeleteTargetWorksheet(LinksSheetName);
            wb.DeleteTargetWorksheet(FilesSheetName);
            wb.DeleteTargetWorksheet(ErrorsSheetName);

            wb.WriteLinks(new ExternalLinks(wb?.Application, nameList));
        }

        public const string LinksSheetName  = "Links Analysis";
        public const string FilesSheetName  = "Linked Files";
        public const string ErrorsSheetName = "Links Errors";

        [SuppressMessage( "Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)" )]
        internal static void WriteLinks(this Workbook wb, IExternalLinks links) {
            if(links.Count == 0  && links.Errors.Count == 0) {
                MessageBox.Show("No external links found!", "", MessageBoxButtons.OK,MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1);
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
            foreach(var columnHeader in columnHeaders) { ws.Cells[2,++colNo].Value = columnHeader; }

            ws.Range[ws.Cells[2,1], ws.Cells[lastRow,columnHeaders.Count]].Columns.AutoFit();

            ws.Range["A1:D1"].Merge();
            ws.Range["A1"].Formula = $"Link analysis run on {DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}";

            ws.Range["$A$2", ws.Cells[lastRow, columnHeaders.Count]].AutoFilter();
            ws.Select();
            ws.Application.ActiveWindow.FreezePanes = false;
            ws.Application.ActiveWindow.SplitColumn = 0;
            ws.Application.ActiveWindow.SplitRow = 2;
            ws.Application.ActiveWindow.FreezePanes = true;
        }

        private static void DeleteTargetWorksheet(this Workbook wb, string sheetName) {
            try {
                wb.Application.DisplayAlerts = false;
                wb.Worksheets[sheetName].Delete();
            } catch (COMException) {
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