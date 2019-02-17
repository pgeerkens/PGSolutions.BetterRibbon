////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.IO;

using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;
    using Range = Microsoft.Office.Interop.Excel.Range;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public static partial class ExcelExtensions {
        internal static SourceCellRef NewCellRef(this Worksheet ws, Range cl) =>
            new SourceCellRef(ws.Parent.Path, ws.Parent.Name, ws.Name, cl.Address);

        internal static SourceCellRef NewWorkbookNameRef(this Workbook wb, Excel.Name namedRange) {
            string sheetName = (namedRange.Parent == wb)
                             ? wb.Name
                             : $"[{namedRange.Parent.name}]";
            return new SourceCellRef(wb.Path, wb.Name, sheetName,
                namedRange.Name.Replace($"'{sheetName}'!", "").Replace($"{sheetName}!", ""));
        }

        public static int RowsCount(this IReadOnlyList<ICellRef> list) => list?.Count ?? 0;
        public static int ColsCount(this IReadOnlyList<ICellRef> list) => list==null ? 0 : 12;
        public static string Item(this IReadOnlyList<ICellRef> list, int row, int col) {
            if (list==null) throw new ArgumentNullException(nameof(list));

            switch (col) {
                case 0: return Path.Combine(list[row].TargetPath, list[row].TargetFile);
                case 1: return list[row].TargetPath;
                case 2: return list[row].TargetFile;
                case 3: return list[row].TargetTab;
                case 4: return list[row].TargetCell;
                case 5: return list[row].IsNamedRange ? "Named Range" : "Cell";
                case 6: return Path.Combine(list[row].SourcePath, list[row].SourceFile);
                case 7: return list[row].SourcePath;
                case 8: return list[row].SourceFile;
                case 9: return list[row].SourceTab;
                case 10: return list[row].SourceCell;
                case 11: return $"'{list[row].Formula}";
                default: throw new ArgumentOutOfRangeException($"Column index {col} out of bounds.");
            }
        }
    }
}
