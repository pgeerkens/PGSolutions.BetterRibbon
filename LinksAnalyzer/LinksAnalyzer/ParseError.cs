////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using PGSolutions.RibbonUtilities.LinksAnalyzer.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IParseError))]
    public class ParseError : IParseError {
        public ParseError(ISourceCellRef cellRef, string formula, int charPosition, string condition) {
            CellRef      = cellRef;
            Formula      = formula;
            Condition    = condition;
            CharPosition = charPosition;
        }

        public ISourceCellRef CellRef {  get; }
        public string Formula      { get; }
        public int    CharPosition { get; }
        public string Condition    {  get; }
    }
}
