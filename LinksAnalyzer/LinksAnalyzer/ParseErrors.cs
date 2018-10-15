////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IParseErrors))]
    public class ParseErrors : IParseErrors, IReadOnlyList<IParseError> { 
        public ParseErrors() => Errors = new List<IParseError>();

        public int          Count => Errors.Count;
        public IParseError  this[int index] => Errors[index];

        private List<IParseError> Errors { get; }

        public void Add(IParseError parseError) => Errors.Add(parseError);

        public void AddFileAccessError(string fullPath, string action) {
            var cellRef = new SourceCellRef(fullPath, "", "", "");
            Errors.Add(new ParseError(cellRef, fullPath, 0, "File not found"));
        }

        public IEnumerator<IParseError> GetEnumerator() => ((IReadOnlyList<IParseError>)Errors).GetEnumerator();
                IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<IParseError>)Errors).GetEnumerator();
    }
}
