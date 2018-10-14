using System.Collections.Generic;

namespace ExcelRibbon.LinksAnalyzer {
    public class ParseErrors { //: IList<ParseError> {
        public ParseErrors() {
            Errors = new List<ParseError>();
        }

        List<ParseError> Errors { get; }

        public Token Add(ParseError parseError) {
            Errors.Add(parseError);
            return Token.ScanError;
        }

        public void AddFileAccessError(string fullPath, string action) {
            var cellRef = new InternalCellRef(fullPath, "", "", "");
            Add(new ParseError(cellRef, fullPath, 0, "File not found"));
        }
    }
}
