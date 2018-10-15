////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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

        /// <summary>Attaches an {IIntegerSource} to the specified DropDown control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description( "Attaches an {IIntegerSource} to the specified DropDown control." )]
        [DispId( 2 )]
        ISourceCellRef NewSourceCellRef(Excel.Workbook wkbk, string tabName, string cellName);

        ISourceCellRef NewSourceCellRef2(string wkBkPath, string wkBkName, string tabName, string cellName,
            bool isNamedRange = false);

        IExternalLinks NewExternalLinks(Excel.Application excel, VBA.Collection nameList);

        IExternalLinks NewExternalLinksWB(Excel.Workbook wb, string excludedName);

        IExternalLinks NewExternalLinksWS(Excel.Worksheet ws);

        IExternalLinks Parse(ISourceCellRef cellRef, string formula);
    }
}
