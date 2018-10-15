////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalFiles))]
    public class ExternalFiles : IExternalFiles, IReadOnlyList<string> {
        public ExternalFiles() => List = new List<string>();

        private List<string> List { get; }

        public int Count => List.Count;
        public string this[int index] => List[index];

        internal ExternalFiles Add(string fileName) {
            List.Add(fileName); return this;
        }

        public IEnumerator<string> GetEnumerator() => ((IReadOnlyList<string>)List).GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<string>)List).GetEnumerator();
    }
}
