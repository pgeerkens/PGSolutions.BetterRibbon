////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Excel = Microsoft.Office.Interop.Excel;
    using Range = Microsoft.Office.Interop.Excel.Range;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>TODO</summary>
    [CLSCompliant(false)]
        public abstract class AbstractParser: AbstractLinksAnalysis, IParser {
        protected AbstractParser() { }

        public event EventHandler<EventArgs<string>> StatusAvailable;

        public ILinksAnalysis Parse() => FuncParse();

        protected abstract  Func<ILinksAnalysis> FuncParse { get; }

        protected ILinksAnalysis ExtendFromWorkbook(Workbook wb) => ExtendFromWorkbook(wb, ExcludedSheetNames);

        protected ILinksAnalysis ExtendFromWorkbook(Workbook wb, IList<string> excludedSheetNames) {
            if (wb == null) throw new ArgumentNullException(nameof(wb));

            foreach(Worksheet ws in wb.Worksheets) {
                if (excludedSheetNames.FirstOrDefault(s => s.Equals(ws.Name)) == null) {
                    ExtendFromWorksheet(ws);
                }
            }

            ExtendFromNamedRanges(wb);
            return this;
        }

        protected ILinksAnalysis ExtendFromWorksheet(Worksheet ws) {
            if (ws == null) return null;

            var usedRange = ws.UsedRange;
            for(var colNo=1; colNo <= usedRange.Columns.Count; colNo++) {
                var percentage = 100 * colNo / usedRange.Columns.Count;
                StatusAvailable?.Invoke(this, 
                    new EventArgs<string>($"Searching {ws.Parent.Name}[{ws.Name}] ... ({percentage,3}%)"));

                var lastRowNo = ws.Cells[ws.Rows.Count, colNo].End(Excel.XlDirection.xlUp).Row;
                for(long rowNo = 1; rowNo <= lastRowNo; rowNo++) {
                    var cell    = usedRange[rowNo, colNo];
                    if ( cell.Formula is string formula && formula.Length > 0 && formula[0] == '=' ) {
                        var cellRef = ws.NewCellRef(cell as Range);
                        ParseFormula(cellRef,formula);
                    }
                }
            }
            return this;
        }

        protected ILinksAnalysis ExtendFromWorkbookList(Range range) {
            if (range==null) return null;

            StatusAvailable?.Invoke(this, new EventArgs<string>("Loading background processor ..."));
            var nameList = range.GetNameList();
            using (var newExcel = WorkbookProcessor.New(range.Application, true)) {
                foreach (var item in nameList) {
                    if (item is string path) {
                        if (!File.Exists(path)) {
                            AddFileAccessError(path, "File not found.");
                            continue;
                        }

                        StatusAvailable?.Invoke(this, new EventArgs<string>($"Processing {path} ..."));

                        try {
                            newExcel.DoOnWorkbook(item, wb=>ExtendFromWorkbook(wb));
                        }
                        catch (IOException ex) { AddFileAccessError(path, $"IOException: '{ex.Message}'"); }
                        finally {
                            StatusAvailable?.Invoke(this, new EventArgs<string>("Ready"));
                        }
                    }
                }
            }
            return this;
        }

        private void ExtendFromNamedRanges(Workbook wb) {
            foreach (Excel.Name source in wb.Names) {
                if (source.RefersTo is string formula  &&  formula.Length > 0
                &&  formula[0] == '=') {
                    var cellRef = wb.NewWorkbookNameRef(source);
                    ParseFormula(cellRef, formula);
                }
            }
        }

        /// <inheritdoc/>
        public const string LinksSheetName  = "Links Analysis";
        /// <inheritdoc/>
        public const string FilesSheetName  = "Linked Files";
        /// <inheritdoc/>
        public const string ErrorsSheetName = "Links Errors";

        static IList<string> ExcludedSheetNames => new List<string> {
            LinksSheetName, FilesSheetName, ErrorsSheetName
        };
    }
}
