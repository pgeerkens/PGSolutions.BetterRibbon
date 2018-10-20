////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    //public class LinksAnalyzer {
    //    public static ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula) => 
    //        new LinksLexer(cellRef, formula);

    //    [CLSCompliant(false)]
    //    public static void ListExternalLinksActiveWorkbook(Excel.Workbook wb, bool IncludeHyperLinks) {

    //        if( wb == null) return;

    //        wb.Application.ScreenUpdating = false;
    //        bool protectStructure = wb.ProtectStructure;
    //        try {
    //            if (wb.ProtectStructure) wb.Protect(null, false);

    //            var externalLinks = new ExternalLinks(wb, "Links Analysis");
                
    //        } finally {
    //            wb.Protect(null, protectStructure);
    //            wb.Application.ScreenUpdating = true;
    //        }
    //    }
    //}
}
