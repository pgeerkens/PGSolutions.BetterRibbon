////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    /// <summary>.</summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksAnalysis))]
    [Guid(Guids.AbstractLinksParser)]
    public class LinksAnalysis : ILinksAnalysis {
        public LinksAnalysis() {
            _errors = new ParseErrors();
            _files  = new FilesDictionary();
            _links  = new ExternalLinks();
        }

        /// <inheritdoc/>
        public IParseErrors   Errors => _errors;
        /// <inheritdoc/>
        public IExternalFiles Files  => _files.OrderedList;
        /// <inheritdoc/>
        public IExternalLinks Links  => _links;

        private readonly ParseErrors     _errors;
        private readonly FilesDictionary _files;
        private readonly ExternalLinks   _links;

        private bool Add(ICellRef cell) {
            if (cell != null) {
                _links.Add(cell);
                _files.Add(Path.Combine(cell.TargetPath, cell.TargetFile));
            }
            return cell != null;
        }

        protected void AddFileAccessError(string path, string condition)
        => _errors.AddFileAccessError(path, condition);

        private void AddParseError(ISourceCellRef cellRef, string formula, int charPosition, string condition)
        =>_errors.Add(new ParseError(cellRef,formula,charPosition, condition));

        [CLSCompliant(false)]
        public ILinksAnalysis ParseFormula(ISourceCellRef sourceCell, string formula) {
            var lexer = new LinksLexer(sourceCell, formula);

            for (var token = lexer.Scan(); token.Value != EToken.EOT; token = lexer.Scan()) {
                switch (token.Value) {
                    case EToken.ScanError:
                        AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Scan error at position {lexer.CharPosition}; found: '{token.Text}'");
                        break;
                    case EToken.ExternRef:
                        var path = token.Text;
                        if((token = lexer.Scan()).Value != EToken.Bang) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected '!' found '{token.Name()}' at position {lexer.CharPosition}");
                        } else if((token = lexer.Scan()).Value != EToken.Identifier) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected Identifier, found '{token.Name()}' at position {lexer.CharPosition}");
                        } else if (! ParseExternRef(path,token.Text,formula,sourceCell)) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected a cell reference at position {lexer.CharPosition}; found '{token.Text}'");
                        } else {
                            break;
                        }
                        break;
                    case EToken.OpenExternRef:
                        path = token.Text;
                        if((token = lexer.Scan()).Value != EToken.Bang) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected '!' found '{token.Name()}' at position {lexer.CharPosition}");
                        } else if((token = lexer.Scan()).Value != EToken.Identifier) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected Identifier, found '{token.Name()}' at position {lexer.CharPosition}");
                        } else if (! ParseOpenExternRef(path,token.Text,formula,sourceCell)) {
                            AddParseError(sourceCell, lexer.Formula, lexer.CharPosition,
                                $"Expected a cell reference at position {lexer.CharPosition}; found '{token.Text}'");
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
            var indexBra  = path.IndexOf('[',       0); if (indexBra < 0) return IsValidSheetName(path);
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
            var indexKet  = path.IndexOf(']',0); if (indexKet < 0) return IsValidSheetName(path);
            return Add(new ExternalRef(formula,source,
                       new SourceCellRef(
                           "open workbook w/o a path",
                           path.Substring(         1, indexKet - 1),               // omit "'['
                           path.Substring(indexKet+1, path.Length - indexKet - 1), // omit ']' trailing
                           cell
            ) ) );
        }

        private static bool IsValidSheetName(string path) {
            var invalid = new List<char>{':', '\\', '/', '?', '*', '[', ']' };
            foreach (var c in invalid) if (path.IndexOf(c) >= 0) return false;

            if (path.IndexOf('\'') == 0  &&  path.Substring(1).IndexOf('\'') == path.Length-2) return true;
            return path.IndexOf('\'') < 0;
        }
    }
}
