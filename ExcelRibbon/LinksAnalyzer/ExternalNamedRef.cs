using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelRibbon.LinksAnalyzer {
    public class ExternalNamedRef {
        /// <summary>Parses a InternalCellRef source, returning a {ExternalNamedRef}</summary>
        /// <param name="path"></param>
        /// <param name="textIn"></param>
        /// <param name="source"></param>
        /// <param name="sourceFile"></param>
        /// <param name="formula"></param>
        /// <returns>Null if the parse fails; else a fully formed {ExternalNamedRef}.</returns>
        public static ExternalNamedRef Parse(string path, string textIn, string formula,
            InternalCellRef source
        ) {
            var text = $"{path}!{textIn}";

            if (text[1] != '\'') return null;
            var skip = 1;
            var indexBra  = text.IndexOf('[', skip);    if (indexBra < 0) return null;
            var indexKet  = text.IndexOf(']',indexBra); if (indexKet < 0) return null;
            var indexBang = text.IndexOf('!',indexKet); if (indexBang < 0) return null;
                
            return new ExternalNamedRef(formula,source,
                new InternalCellRef(
                text.Substring(2, indexBra - 2),
                text.Substring(indexBra+1, indexKet  - indexBra - 1),
                text.Substring(indexKet+1, indexBang - indexKet - 1),
                text.Substring(indexBang+1,text.Length - indexBang-1))
            );
        }

        /// <summary>Parses a NamedRange source, returning a {ExternalNamedRef}</summary>
        /// <param name="path"></param>
        /// <param name="textIn"></param>
        /// <param name="source"></param>
        /// <param name="sourceFile"></param>
        /// <param name="formula"></param>
        /// <returns>Null if the parse fails; else a fully formed {ExternalNamedRef}.</returns>
        public static ExternalNamedRef Parse(string path, string textIn, string formula,
                Excel.Name source, string sourceFile
        ) {
            return Parse(path, textIn, formula,
                new InternalCellRef(sourceFile,source.Parent,source.Name,formula,true));
        }

        private ExternalNamedRef(string formula, InternalCellRef source, InternalCellRef target) {
            Formula = formula;
            Source = source;
            Target = target;
        }

        public  string          Formula { get; }
        private InternalCellRef Source  { get; }
        private InternalCellRef Target  { get; }

        public string TargetPath     => Target.FullPath;
        public string TargetFileName => Target.FileName;
        public string TargetTabName  => Target.TabName;
        public string TargetCellName => Target.CellName;

        public string SourcePath     => Source.FullPath;
        public string SourceFileName => Source.FileName;
        public string SourceTabName  => Source.TabName;
        public string SourceCellName => Source.CellName;
    }
}
