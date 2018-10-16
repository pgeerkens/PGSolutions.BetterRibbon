////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    [Serializable]
    public class FilesDictionary : Dictionary<string,string> {
        public FilesDictionary() : base() { }

        internal void Add(string fileName) {
            if(!Keys.Contains(fileName)) { Add(fileName,fileName); }
        }

        internal ExternalFiles OrderedList {
            get {
                var files = new ExternalFiles();
                foreach(var file in this.OrderBy(i=>i.Key)) files.Add(file.Value);
                return files;
            }
        }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalFiles))]
    public class ExternalFiles : IExternalFiles{ //, IReadOnlyList<string> {
        public ExternalFiles() => List = new List<string>();

        public int      Count => List.Count;
        public string   Item(int index) => List[index];
//        public string this[int index] => List.Keys.OrderBy(i=>i).Skip(index).Take(1).FirstOrDefault();

        private List<string> List { get; }

        internal void Add(string fileName) => List.Add(fileName);

//        public IEnumerator<string> GetEnumerator() => ((IReadOnlyList<string>)List).GetEnumerator();
//        IEnumerator IEnumerable.GetEnumerator() => ((IReadOnlyList<string>)List).GetEnumerator();
    }
}
