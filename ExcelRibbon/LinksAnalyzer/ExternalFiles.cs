using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRibbon.LinksAnalyzer {
    public class ExternalFiles {
        public ExternalFiles() => List = new List<string>();

        private List<string> List { get; }

        public int Count => List.Count;
        public string this[int index] => List[index];

        internal ExternalFiles Add(string fileName) {
            List.Add(fileName); return this;
        }
    }
}
