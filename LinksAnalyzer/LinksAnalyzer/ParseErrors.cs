////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [Description("")]
    [Serializable]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IParseErrors))]
    public class ParseErrors : IParseErrors, IReadOnlyList<IParseError> { 
        internal ParseErrors() => Errors = new List<IParseError>();

        public int          Count => Errors.Count;
        public IParseError  this[int index] => Errors[index];

        private List<IParseError> Errors { get; }

        public void Add(IParseError parseError) => Errors.Add(parseError);

        public void AddFileAccessError(string fullPath, string condition) {
            var cellRef = new SourceCellRef(fullPath, "", "", "");
            Errors.Add(new ParseError(cellRef, fullPath, 0, condition));
        }

        public IEnumerator<IParseError> GetEnumerator() => ((IReadOnlyList<IParseError>)Errors).GetEnumerator();
                IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<IParseError>)Errors).GetEnumerator();
    }
}
