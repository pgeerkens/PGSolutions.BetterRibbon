using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRibbon.LinksAnalyzer {
    public class ParseError {
        public ParseError(InternalCellRef cellRef, string formula, long charPosition, string condition) {
            CellRef = cellRef;
            Formula = formula;
            Condition = condition;
            CharPosition = charPosition;
        }

        public InternalCellRef CellRef {  get; }
        public string Formula      { get; }
        public long   CharPosition { get; }
        public string Condition    {  get; }
    }
}
