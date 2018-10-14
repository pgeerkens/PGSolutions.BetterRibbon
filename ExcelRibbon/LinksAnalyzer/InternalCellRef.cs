namespace ExcelRibbon.LinksAnalyzer {
    public class InternalCellRef {
        public InternalCellRef(string wkBkPath, string wkBkName, string tabName, string cellName,
            bool isNamedRangeRef = false
        ) {
            FullPath = wkBkPath;
            FileName = wkBkName;
            TabName  = tabName;
            CellName = cellName;
            IsNamedRangeCellRef = IsNamedRangeCellRef;
        }

        public bool   IsNamedRangeCellRef { get;}
        public string CellName { get; }
        public string TabName  { get; }
        public string FileName {get; }
        public string FullPath { get; }
    }
}
