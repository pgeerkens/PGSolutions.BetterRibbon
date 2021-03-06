﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using static AbstractParser;

    /// <summary>Extension methods for Excel objects.</summary>
    [CLSCompliant(false)]
    public static class ExcelLinksExtensions {
        ///// <inheritdoc/>
        //public static ImageObject ToggleImage(this bool isPressed)
        //=> isPressed ? "TagMarkComplete" : "MarginsShowHide";

        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public static void WriteLinks(this Workbook wb, ILinksAnalysis links) {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (links == null) throw new ArgumentNullException(nameof(links));
            
            wb.DeleteTargetWorksheet(LinksSheetName);
            wb.DeleteTargetWorksheet(FilesSheetName);
            wb.DeleteTargetWorksheet(ErrorsSheetName);

            if (links.Links.Count == 0  && links.Errors.Count == 0) {
                MessageBox.Show("No external links found!", "", MessageBoxButtons.OK, MessageBoxIcon.Information,
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

        public static void WriteLinksAnalysis(this Worksheet ws, ILinksAnalysis links) {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (links == null) throw new ArgumentNullException(nameof(links));

            if (links.Links.Count > 0) {
                var calculation = ws.Application.Calculation;
                try {
                    ws.Application.Calculation = XlCalculation.xlCalculationManual;
                    ws.Application.ScreenUpdating = false;

                    var links2D   = links.Links;
                    var MaxCol    = links2D.ColsCount();
                    var firstCell = ws.Cells[3,1];
                    var lastCell  = ws.Cells[links2D.RowsCount() + 2, MaxCol];
                    var sheetData = ws.Range[firstCell,lastCell] as Range;

                    ws.Columns[MaxCol].EntireColumn.NumberFormat = "@"; // Formula column
                    links2D.FastCopyToRange(sheetData);

                    ws.InitializeTargetWorksheet(links2D.RowsCount() + 2, new List<string>() {
                            "Target FullName", "Target Path", "Target FileName", "Target Worksheet", "Target Cell", "Link Type",
                            "Source FullName", "Source Path", "Source FileName", "Source Worksheet", "Source Cell", "Source Formula"});
                }
                finally {
                    ws.Application.ScreenUpdating = true;
                    ws.Application.Calculation = calculation;
                }

            }
        }

        public static void WriteLinksFiles(this Worksheet ws, IExternalFiles files) {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (files == null) throw new ArgumentNullException(nameof(files));

            var lastRow = 2;
            var i       = 0;
            foreach (var fileName in files.OrderBy(s => s)) {
                i++; lastRow++;
                ws.Cells[lastRow, 1].Value2 = fileName;

                ws.WritePercentageStatus(ws.Name, 100*i/files.Count);
            }
            ws.InitializeTargetWorksheet(lastRow, new List<string>() { "External FIles" });
        }

        public static void WriteLinksErrors(this Worksheet ws, IParseErrors errors) {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (errors == null) throw new ArgumentNullException(nameof(errors));

            var lastRow = 2;
            var i       = 0;
            foreach (var error in errors) {
                i++; lastRow++;
                var col = 0;
                ws.Cells[lastRow, ++col].Value2 = error.CellRef.CellName;
                ws.Cells[lastRow, ++col].Value2 = error.CellRef.TabName;
                ws.Cells[lastRow, ++col].Value2 = error.CellRef.FileName;
                ws.Cells[lastRow, ++col].Value2 = error.Condition; ;
                ws.Cells[lastRow, ++col].Value2 = error.CharPosition;
                ws.Cells[lastRow, ++col].Value2 = $"'{error.Formula}";
                ws.Cells[lastRow, ++col].Value2 = error.CellRef.FullPath;

                ws.WritePercentageStatus(ws.Name, 100*i/errors.Count);
            }

            ws.InitializeTargetWorksheet(lastRow, new List<string>() {
                    "Cell Name", "Tab Name", "File Name", "Error Condition", "Position", "Formula", "Path"});
        }

        private static void WritePercentageStatus(this Worksheet ws, string sheetName, int percentage)
            => ws.Application.StatusBar = $"Writing worksheet {sheetName}: ({percentage}%) ... ";

        private static void InitializeTargetWorksheet(this Worksheet ws, long lastRow, IList<string> columnHeaders) {
            var colNo = 0;
            foreach (var columnHeader in columnHeaders) { ws.Cells[2, ++colNo].Value = columnHeader; }

            ws.Range[ws.Cells[2, 1], ws.Cells[lastRow, columnHeaders.Count]].Columns.AutoFit();

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
            }
            catch (COMException) { /*  NO-OP: Sheet doesn't exist so delete not needed */ }
            finally {
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
