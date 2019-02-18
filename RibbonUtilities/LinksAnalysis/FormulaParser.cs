////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Range = Microsoft.Office.Interop.Excel.Range;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    [CLSCompliant(false)]
    public sealed class FormulaParser {
        /// <summary>Returns all the external links found in the supplied formula.</summary>
        public FormulaParser() : base() {
            LinksParser = new LinksAnalysis();
            LinksParser.StatusAvailable += OnStatusAvailable;
        }

        private LinksAnalysis LinksParser { get; }

        public event EventHandler<EventArgs<string>> StatusAvailable;

        private void OnStatusAvailable(object sender, EventArgs<string> e)
        => StatusAvailable?.Invoke(sender,e);

        /// <summary>Returns all the external links found in the supplied formula.</summary>
        public ILinksAnalysis ParseFormula(ISourceCellRef cellRef, string formula)
        => LinksParser.ParseFormula(cellRef, formula);

        /// <summary>Returns all the external links found in the supplied {Excel.Worksheet}.</summary>
        public ILinksAnalysis ParseWorksheet(Worksheet ws)
        => LinksParser.ExtendFromWorksheet(ws);

        /// <summary>Returns all the external links found in the supplied {Excel.Workbook}.</summary>
        public ILinksAnalysis ParseWorkbook(Workbook wb)
        => LinksParser.ExtendFromWorkbook(wb);

        /// <summary>Returns all the external links found in the supplied {Excel.Workbook}.</summary>
        public ILinksAnalysis ParseWorkbook(Workbook wb, IList<string> excludedSheetNames)
        => LinksParser.ExtendFromWorkbook(wb, excludedSheetNames);

        /// <summary>Returns all the external links found in the supplied list of workbook names.</summary>
        public ILinksAnalysis ParseWorkbookList(Range range, bool inBackground)
        => LinksParser.ExtendFromWorkbookList(range, inBackground);
    }
}
