////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalLinks))]
    public sealed class ExternalLinks : IExternalLinks, IReadOnlyList<ICellRef> {
        public ExternalLinks(ISourceCellRef cellRef, string formula) : this() => ParseFormula(cellRef, formula);
        public ExternalLinks(Excel.Worksheet ws) : this() => ExtendFromWorksheet(ws);
        public ExternalLinks(Excel.Workbook wb, string excludedName) : this() {
            if (wb == null) return;

            foreach(Excel.Worksheet ws in wb.Worksheets) {
                if ( ! excludedName.Equals(ws.Name) ) { ExtendFromWorksheet(ws); }
                // DoEvents
            }

            ExtendFromNamedRanges(wb);
        }
        private ExternalLinks() {
            Links   = new List<ICellRef>();
            Errors  = new ParseErrors();
        }

        public   int             Count => Links.Count;
        public   ICellRef        this[int index] => Links[index];
        private  IList<ICellRef> Links;

        internal ParseErrors     Errors   { get; }
        internal ExternalFiles   Files    { get; private set; }

        public IEnumerator<ICellRef> GetEnumerator() => ((IReadOnlyList<ICellRef>)Links).GetEnumerator();
             IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<ICellRef>)Links).GetEnumerator();

        private void ExtendFromWorksheet(Excel.Worksheet ws) {
            if (ws == null) return;

            var messageText = $"Searching {ws.Parent.Name}[{ws.Name}] ... (???%)";
            var usedRange = ws.UsedRange;
            for(var colNo=1; colNo <= usedRange.Columns.Count; colNo++) {
                var percentage = 100 * colNo / usedRange.Columns.Count;
                ws.Application.StatusBar = messageText.Replace("???", percentage.ToString().PadLeft(3));

                var lastRowNo = ws.Cells[ws.Rows.Count, colNo].End(Excel.XlDirection.xlUp).Row;
                for(long rowNo = 1; rowNo <= lastRowNo; rowNo++) {
                    var cell    = ws.Cells[rowNo, colNo];
                    if ( cell.Formula is string formula && formula.Length > 0 && formula[0] == '=' ) {
                        var cellRef = NewCellRef(ws,cell);
                        ParseFormula(cellRef,formula);
                    }
                    // DoEvents
                }
            }
        }

        private void ExtendFromNamedRanges(Excel.Workbook wb) {
            foreach(Excel.Name source in wb.Names) {
                if ( source.RefersTo is string formula  &&  formula.Length > 0  
                &&  formula[0] == '=') {
                    var cellRef = NewWorkbookNameRef(wb, source);
                    ParseFormula(cellRef,formula);
                }
                // DoEvents
            }
        }

        private SourceCellRef NewCellRef(Excel.Worksheet ws, Excel.Range cl) =>
            new SourceCellRef(ws.Parent.Path, ws.Parent.Name, ws.Name, cl.Address);

        private SourceCellRef NewWorkbookNameRef(Excel.Workbook wb, Excel.Name namedRange) {
            string sheetName = (namedRange.Parent == wb)
                             ? wb.Name
                             : $"[{namedRange.Parent.name}]";
            return new SourceCellRef(wb.Path, wb.Name, sheetName,
                namedRange.Name.Replace($"'{sheetName}'!", "").Replace($"{sheetName}!", ""));
        }

        private IExternalLinks ParseFormula(ISourceCellRef sourceCell, string formula) {
            var lexer = new LinksLexer(sourceCell, formula);

            for (var token = lexer.Scan(); token.Value != EToken.EOT; token = lexer.Scan()) {
                switch (token.Value) {
                    case EToken.ScanError:
                        Errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Scan error at position {lexer.CharPosition}; found: '{token.Text}'"));
                        break;
                    case EToken.ExternRef:
                        var path = token.Text;
                        if((token = lexer.Scan()).Value != EToken.Bang) {
                            Errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected '!' found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if((token = lexer.Scan()).Value != EToken.Identifier) {
                            Errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected Identifier, found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if (! ParseExternRef(path,token.Text,formula,sourceCell)) {
                            Errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected a cell reference at position {lexer.CharPosition}; found '{token.Text}'"));
                        } else {
                            break;
                        }
                        break;
                    default:
                        break;
                }
            }
            return this;
        }

        private bool ParseExternRef(string path, string cell, string formula, ISourceCellRef source) {
            var indexBra  = path.IndexOf('[',       0); if (indexBra < 0) return false;
            var indexKet  = path.IndexOf(']',indexBra); if (indexKet < 0) return false;
            return Add(new ExternalRef(formula,source,
                       new SourceCellRef(
                           path.Substring(         1, indexBra - 1),               // omoi "'" leading
                           path.Substring(indexBra+1, indexKet - indexBra - 1),    // omit "'['
                           path.Substring(indexKet+1, path.Length - indexKet - 2), // omit ']' trailing
                           cell
            ) ) );
        }

        private bool Add(ICellRef cell) {
            if(cell != null) { Links.Add(cell); }
            return cell != null;
        }
    }
}
