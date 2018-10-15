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
        public SourceCellRef(Excel.Workbook wkbk, string tabName, string cellName) 
            : this(wkbk.Path, wkbk.Name, tabName, cellName) { }
        public SourceCellRef(string wkBkPath, string wkBkName, string tabName, string cellName,
            bool isNamedRange = false
        ) {
            FullPath = wkBkPath;
            FileName = wkBkName;
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
