////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook  = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ILinksAnalyzer)]
    public interface ILinksAnalyzer {
        /// <summary>Attaches an {IIntegerSource} to the specified DropDown control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description( "Attaches an {IIntegerSource} to the specified DropDown control." )]
        [DispId(1)]
        ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula);

        IExternalLinks NewExternalLinks(Excel.Application excel, INameList nameList);

        IExternalLinks NewExternalLinksWB(Workbook wb, string excludedName);

        IExternalLinks NewExternalLinksWS(Worksheet ws);

        IExternalLinks Parse(ISourceCellRef cellRef, string formula);

        /// <summary></summary>
        /// <param name="wb"></param>
        void WriteLinksAnalysisWB(Workbook wb);

        /// <summary></summary>
        /// <param name="wb"></param>
        /// <param name="nameList"></param>
        void WriteLinksAnalysisFiles(Workbook wb, INameList nameList);
    }
}
