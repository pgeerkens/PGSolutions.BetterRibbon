////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    public class LinksAnalyzer {
        public static ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula) => 
            new LinksLexer(cellRef, formula);

        [CLSCompliant(false)]
        public static void ListExternalLinksActiveWorkbook(Excel.Workbook workbook,
            bool showErrors, bool IncludeHyperLinks) {

            if( workbook == null) return;

            workbook.Application.ScreenUpdating = false;
            bool protectStructure = workbook.ProtectStructure;
            try {
                if (workbook.ProtectStructure) workbook.Protect(null, false);

                var externalLinks = new ExternalLinks(workbook, "Links Analysis");
                
            } finally {
                workbook.Protect(null, protectStructure);
                workbook.Application.ScreenUpdating = true;
            }
        }
    }
}
