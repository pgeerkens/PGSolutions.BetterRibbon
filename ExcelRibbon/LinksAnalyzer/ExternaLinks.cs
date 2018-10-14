using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelRibbon.LinksAnalyzer {
    public class ExternalLinks : ICellRef {
        public ExternalLinks() {
            List = new List<ICellRef>();
        }

        public   ParseErrors     Errors   { get; }
        private  InternalCellRef Location { get; set; }
        private  LinksLexer      Lexer    { get; }
        internal ExternalFiles   Files    { get; private set; }

        private IList<ICellRef> List;

        public int      Count => List.Count;
        public ICellRef this[int index] => List[index];


        public ExternalLinks Parse(LinksLexer lexer, InternalCellRef cellRef) {
            Location = cellRef;

            for (var token = lexer.Scan(); token.Token != Token.EOT; token = lexer.Scan()) {
                switch (token.Token) {
                    case Token.ScanError:
                        Errors.Add(lexer.RaiseError(cellRef, "Unknown token found"));
                        break;
                    case Token.ExternRef:
                        ParseExternRef(lexer, token.Text, cellRef);
                        return null;
                    default:
                        break;
                }
            }
            return this;
        }

        private InternalCellRef NewCellRef(Excel.Worksheet ws, Excel.Range cl) =>
            new InternalCellRef(ws.Parent.Path, ws.Parent.Name, ws.Name, cl.Address);

        private InternalCellRef NewWorkbookNameRef(Excel.Workbook wb, Excel.Name namedRange) {
            string sheetName = (namedRange.Parent == wb)
                             ? "<workbook>"
                             : namedRange.Parent.name;
            return new InternalCellRef(wb.Path, wb.Name, sheetName,
                namedRange.Name.Replace($"'{sheetName}'!", "").Replace($"{sheetName}!", ""));
        }
        private void ParseExternRef(LinksLexer lexer, string path, InternalCellRef cellRef) {
            var token = lexer.Scan();
            if (token.Token == Token.Bang) token = lexer.Scan();
            if (token.Token != Token.Identifier) {
                return Errors.Add(lexer.RaiseError(Location, "Expected identifier"));
            } else {
                if (cellRef.IsNamedRangeCellRef) {
                    var er = new ExternalNamedRef();
                    if (er.Parse(path, token.Text, cellRef, lexer.TextIn)) {
                        List.Add(this);
                        Files.Add(this.ICellRef_Path + this.ICellRef_FileName);
                    }
                } else {
                    var er = new ExternalNamedRef();
                    List.Add(this);
                    Files.Add(this.ICellRef_Path + this.ICellRef_FileName);
                }
            }
        }

    }

    public interface ICellRef {

    }
}
