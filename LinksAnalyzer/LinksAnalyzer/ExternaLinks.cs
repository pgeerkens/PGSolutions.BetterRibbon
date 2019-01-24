////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [Serializable]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalLinks))]
    public sealed class ExternalLinks : IExternalLinks, ITwoDimensionalLookup, IReadOnlyList<ICellRef> {
        /// <summary>Returns all the external links found in the supplied formula.</summary>
        public ExternalLinks(ISourceCellRef cellRef, string formula) : this() 
            => RunAndBuildFiles(() => ParseFormula(cellRef, formula));
        /// <summary>Returns all the external links found in the supplied {Excel.Worksheet}.</summary>
        public ExternalLinks(Excel.Worksheet ws) : this()
            => RunAndBuildFiles(() => ExtendFromWorksheet(ws));
        /// <summary>Returns all the external links found in the supplied {Excel.Workbook}.</summary>
        public ExternalLinks(Excel.Workbook wb, string excludedName) : this()
            => RunAndBuildFiles(() => ExtendFromWorkbook(wb, excludedName));
        /// <summary>Returns all the external links found in the supplied list of workbook names.</summary>
        public ExternalLinks(Excel.Application excel, INameList nameList) : this() {
            if(excel==null)throw new ArgumentNullException("excel","Supplied argument may not be null.");
            if(nameList==null) return;
            for(var i=0; i<nameList.Count; i++) {
                var item = nameList.Item(i);
                if ( item is string path) {
                    if(!File.Exists(path)) {
                        _errors.AddFileAccessError(path,"File not found.");
                        continue;
                    }

                    excel.Application.StatusBar = $"Opening {path}";
                    Excel.Workbook wb = null;
                    try {
                        excel.Application.DisplayAlerts = false;
                        wb = excel.Application.Workbooks.Open(path,UpdateLinks:false,ReadOnly:true,AddToMru:false);

                        ExtendFromWorkbook(wb,"");
                    } catch(IOException ex) {
                        _errors.AddFileAccessError(path,$"IOException: '{ex.Message}'");
                    } finally {
                        wb?.Close(SaveChanges:false);
                        excel.Application.DisplayAlerts = false;
                    }
                    // DoEvents
                }
            }
            Files = _files.OrderedList;
        }
        private ExternalLinks() {
            Links   = new List<ICellRef>();
            _errors = new ParseErrors();
            _files  = new FilesDictionary();
        }
        private void RunAndBuildFiles(Action action) {
            action();
            Files = _files.OrderedList;
        }

        public   int             Count      => Links.Count;
        public   ICellRef        this[int index] => Links[index];
        public   IParseErrors    Errors     => _errors;
        public   IExternalFiles  Files      { get; private set; }

        int     ITwoDimensionalLookup.RowsCount => Count;
        int     ITwoDimensionalLookup.ColsCount => 12;
        string  ITwoDimensionalLookup.Item(int row, int col) {
            switch (col) {
                case  0: return Path.Combine(this[row].TargetPath,this[row].TargetFile);
                case  1: return this[row].TargetPath;
                case  2: return this[row].TargetFile;
                case  3: return this[row].TargetTab;
                case  4: return this[row].TargetCell;
                case  5: return this[row].IsNamedRange ? "Named Range" : "Cell";
                case  6: return Path.Combine(this[row].SourcePath,this[row].SourceFile);
                case  7: return this[row].SourcePath;
                case  8: return this[row].SourceFile;
                case  9: return this[row].SourceTab;
                case 10: return this[row].SourceCell;
                case 11: return $"'{this[row].Formula}";
                default: throw new ArgumentOutOfRangeException($"Column index {col} out of bounds.");
            }
        }

        private  IList<ICellRef> Links;
        private  ParseErrors     _errors    { get; }
        private  FilesDictionary _files     { get; }

        public IEnumerator<ICellRef> GetEnumerator() => ((IReadOnlyList<ICellRef>)Links).GetEnumerator();
             IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<ICellRef>)Links).GetEnumerator();

        private void ExtendFromWorkbook(Excel.Workbook wb, string excludedName) {
            if (wb == null) return;

            foreach(Excel.Worksheet ws in wb.Worksheets) {
                if ( ! excludedName.Equals(ws.Name) ) { ExtendFromWorksheet(ws); }
                // DoEvents
            }

            ExtendFromNamedRanges(wb);
        }

        private void ExtendFromWorksheet(Excel.Worksheet ws) {
            if (ws == null) return;

            var messageText = $"Searching {ws.Parent.Name}[{ws.Name}] ... (???%)";
            var usedRange = ws.UsedRange;
            for(var colNo=1; colNo <= usedRange.Columns.Count; colNo++) {
                var percentage = 100 * colNo / usedRange.Columns.Count;
                ws.Application.StatusBar = messageText.Replace("???", percentage.ToString().PadLeft(3));

                var lastRowNo = ws.Cells[ws.Rows.Count, colNo].End(Excel.XlDirection.xlUp).Row;
                for(long rowNo = 1; rowNo <= lastRowNo; rowNo++) {
                    //var cell    = ws.Cells[rowNo, colNo];
                    var cell    = usedRange[rowNo, colNo];
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

        private static SourceCellRef NewCellRef(Excel.Worksheet ws, Excel.Range cl) =>
            new SourceCellRef(ws.Parent.Path, ws.Parent.Name, ws.Name, cl.Address);

        private static SourceCellRef NewWorkbookNameRef(Excel.Workbook wb, Excel.Name namedRange) {
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
                        _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Scan error at position {lexer.CharPosition}; found: '{token.Text}'"));
                        break;
                    case EToken.ExternRef:
                        var path = token.Text;
                        if((token = lexer.Scan()).Value != EToken.Bang) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected '!' found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if((token = lexer.Scan()).Value != EToken.Identifier) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected Identifier, found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if (! ParseExternRef(path,token.Text,formula,sourceCell)) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected a cell reference at position {lexer.CharPosition}; found '{token.Text}'"));
                        } else {
                            break;
                        }
                        break;
                    case EToken.OpenExternRef:
                        path = token.Text;
                        if((token = lexer.Scan()).Value != EToken.Bang) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected '!' found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if((token = lexer.Scan()).Value != EToken.Identifier) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected Identifier, found '{token.Name()}' at position {lexer.CharPosition}"));
                        } else if (! ParseOpenExternRef(path,token.Text,formula,sourceCell)) {
                            _errors.Add(new ParseError(sourceCell, lexer.Formula, lexer.CharPosition,
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
                           path.Substring(         1, indexBra - 1),               // omit "'" leading
                           path.Substring(indexBra+1, indexKet - indexBra - 1),    // omit "'['
                           path.Substring(indexKet+1, path.Length - indexKet - 2), // omit ']' trailing
                           cell
            ) ) );
        }

        private bool ParseOpenExternRef(string path, string cell, string formula, ISourceCellRef source) {
            var indexKet  = path.IndexOf(']',0); if (indexKet < 0) return false;
            return Add(new ExternalRef(formula,source,
                       new SourceCellRef(
                           "open workbook w/o a path",
                           path.Substring(         1, indexKet - 1),               // omit "'['
                           path.Substring(indexKet+1, path.Length - indexKet - 1), // omit ']' trailing
                           cell
            ) ) );
        }

        private bool Add(ICellRef cell) {
            if(cell != null) {
                Links.Add(cell);
                _files.Add(Path.Combine(cell.TargetPath, cell.TargetFile));
            }
            return cell != null;
        }
    }
}
