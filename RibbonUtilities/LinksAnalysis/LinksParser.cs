////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;

using Core = Microsoft.Office.Core;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Excel = Microsoft.Office.Interop.Excel;
    using Range = Microsoft.Office.Interop.Excel.Range;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>TODO</summary>
    [SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [Serializable]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalysis))]
    public sealed class LinksParser : AbstractLinksParser {
        /// <summary>Returns all the external links found in the supplied formula.</summary>
        public LinksParser(ISourceCellRef cellRef, string formula) : base() 
        => ParseFormula(cellRef, formula);

        /// <summary>Returns all the external links found in the supplied {Excel.Worksheet}.</summary>
        public LinksParser(Worksheet ws) : base()
        => ExtendFromWorksheet(ws);

        /// <summary>Returns all the external links found in the supplied {Excel.Workbook}.</summary>
        public LinksParser(Workbook wb, string excludedName) : base()
        => ExtendFromWorkbook(wb, excludedName);

        /// <summary>Returns all the external links found in the supplied list of workbook names.</summary>
        public LinksParser(Range range) : base()
        => ExtendFromWorkbookList(range);

        public event StatusAvailableEventHandler StatusAvailable;

        private void ExtendFromWorkbook(Workbook wb, string excludedName) {
            foreach(Worksheet ws in wb.Worksheets) {
                if ( ! excludedName.Equals(ws.Name) ) { ExtendFromWorksheet(ws); }
            }

            ExtendFromNamedRanges(wb);
        }

        private void ExtendFromWorksheet(Worksheet ws) {
            if (ws == null) return;

            var messageText = $"Searching {ws.Parent.Name}[{ws.Name}] ... (???%)";
            var usedRange = ws.UsedRange;
            for(var colNo=1; colNo <= usedRange.Columns.Count; colNo++) {
                var percentage = 100 * colNo / usedRange.Columns.Count;
                ws.Application.StatusBar = messageText.Replace("???", percentage.ToString().PadLeft(3));

                var lastRowNo = ws.Cells[ws.Rows.Count, colNo].End(Excel.XlDirection.xlUp).Row;
                for(long rowNo = 1; rowNo <= lastRowNo; rowNo++) {
                    var cell    = usedRange[rowNo, colNo];
                    if ( cell.Formula is string formula && formula.Length > 0 && formula[0] == '=' ) {
                        var cellRef = ws.NewCellRef(cell as Range);
                        ParseFormula(cellRef,formula);
                    }
                }
            }
        }

        private void ExtendFromNamedRanges(Workbook wb) {
            foreach(Excel.Name source in wb.Names) {
                if ( source.RefersTo is string formula  &&  formula.Length > 0  
                &&  formula[0] == '=') {
                    var cellRef = wb.NewWorkbookNameRef(source);
                    ParseFormula(cellRef,formula);
                }
            }
        }

        private void ExtendFromWorkbookList(Range range) {
            if (range==null) return;

            var nameList = range.GetNameList();
            var excel = range.Application;
            var @as = excel.AutomationSecurity;
            try {
                excel.AutomationSecurity = Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;

                foreach (var item in nameList) {
                    if (item is string path) {
                        if (!File.Exists(path)) {
                            AddFileAccessError(path, "File not found.");
                            continue;
                        }

                        excel.ScreenUpdating = true;
                        StatusAvailable?.Invoke(this, $"Processing {path} ....");
                        excel.ScreenUpdating = false;

                        try {
                            var wb = excel.TryItem(item);
                            if (wb == null) {
                                AnalyzeClosedWorkbook(excel, item);
                            } else {
                                ExtendFromWorkbook(wb, "");
                            }
                        }
                        catch (IOException ex) { AddFileAccessError(path, $"IOException: '{ex.Message}'"); }
                    }
                }
            }
            finally {
                excel.ScreenUpdating = true;
                excel.DisplayAlerts = true;
                excel.AutomationSecurity = @as;
            }
        }

        private void AnalyzeClosedWorkbook(Excel.Application excel, string path) {
            Workbook wb = null;
            try {
                wb = excel.Workbooks.Open(path, UpdateLinks: false, ReadOnly: true, AddToMru: false);
                ExtendFromWorkbook(wb, "");
            }
            finally {
                wb?.Close(SaveChanges: false);
            }
        }
    }
}
