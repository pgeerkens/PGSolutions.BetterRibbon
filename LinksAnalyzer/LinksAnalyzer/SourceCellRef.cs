////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISourceCellRef))]
    public class SourceCellRef : ISourceCellRef {
        [CLSCompliant(false)]
        public SourceCellRef(Excel.Workbook workbook, string tabName, string cellName) 
            : this(workbook?.Path, workbook?.Name, tabName, cellName) { }
        [System.Diagnostics.CodeAnalysis.SuppressMessage( "Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed" )]
        public SourceCellRef(string workbookPath, string workbookName, string tabName, string cellName,
            bool isNamedRange = false
        ) {
            FullPath = workbookPath;
            FileName = workbookName;
            TabName  = tabName;
            CellName = cellName;
            IsNamedRange = isNamedRange;
        }

        public bool   IsNamedRange  { get; }
        public string CellName      { get; }
        public string TabName       { get; }
        public string FileName      { get; }
        public string FullPath      { get; }
    }
}
