////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

using PGSolutions.RibbonUtilities.LinksAnalyzer.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer {
    [Serializable]
    [ComVisible(false)]
    public class FilesDictionary : Dictionary<string,string> {
        public FilesDictionary() : base() { }
        protected FilesDictionary(SerializationInfo info, StreamingContext context)
            : base(info,context) { }

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
    [SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [Description("")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IExternalFiles))]
    public class ExternalFiles : IExternalFiles{
        internal ExternalFiles() => List = new List<string>();

        public int    Count           => List.Count;
        public string this[int index] => List[index];

        private List<string> List { get; }

        internal void Add(string fileName) => List.Add(fileName);
        public IEnumerator<string> GetEnumerator() => List.GetEnumerator();
           IEnumerator IEnumerable.GetEnumerator() => List.GetEnumerator();
    }
}
