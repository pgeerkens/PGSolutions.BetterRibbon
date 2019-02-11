////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
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
}
