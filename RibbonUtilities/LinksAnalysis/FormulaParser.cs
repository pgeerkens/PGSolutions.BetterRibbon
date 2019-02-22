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
    /// <summary>Returns all the external links found in the supplied formula.</summary>
    public sealed class FormulaParser: AbstractParser {
        /// <summary>Creates a new <see cref="IParser"/> for a single string formula.</summary>
        public FormulaParser(ISourceCellRef cellRef, string formula)
        => FuncParse = () => ParseFormula(cellRef, formula);

        protected override Func<ILinksAnalysis> FuncParse { get; }
    }

    /// <summary>Returns all the external links found in the supplied <see cref="Worksheet"/>.</summary>
    [CLSCompliant(false)]
    public sealed class WorksheetParser: AbstractParser {
        /// <summary>Creates a new <see cref="IParser"/> for a single <see cref="Worksheet"/>.</summary>
        public WorksheetParser(Worksheet ws)
        => FuncParse = () => ExtendFromWorksheet(ws);

        protected override Func<ILinksAnalysis> FuncParse { get; }
    }

    [CLSCompliant(false)]
    /// <summary>Returns all the external links found in the supplied <see cref="Workbook"/>.</summary>
    public sealed class WorkbookParser: AbstractParser {
        /// <summary>Creates a new <see cref="IParser"/> for a single <see cref="Workbook"/>.</summary>
        /// <param name="wb"></param>
        public WorkbookParser(Workbook wb)
        => FuncParse = () => ExtendFromWorkbook(wb);

        public WorkbookParser(Workbook wb, IList<string> excludedSheetNames)
        => FuncParse = () => ExtendFromWorkbook(wb, excludedSheetNames);

        protected override Func<ILinksAnalysis> FuncParse { get; }
    }

    /// <summary>Returns all the external links found in the <see cref="Workbook"/> names found in <see cref="Range"/>.</summary>
    [CLSCompliant(false)]
    public sealed class WorkbookListParser: AbstractParser {
        /// <summary>Creates a new <see cref="IParser"/> for a list of <see cref="Workbook"/>s.</summary>
        public WorkbookListParser(Range range)
        => FuncParse = () => ExtendFromWorkbookList(range);

        protected override Func<ILinksAnalysis> FuncParse { get; }
    }
}
