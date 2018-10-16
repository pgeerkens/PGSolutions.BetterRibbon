////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    internal static partial class Extensions {
        public static IParseError ErrorHere(this ILinksLexer lexer, string source, string condition) =>
            new ParseError(lexer.CellRef, lexer.Formula, lexer.CharPosition, $"{source}: {condition}");

        /// <summary>Returns true IFF current character is between '0' and '9'.</summary>
        public static bool IsNumeric     (this char c) => '0' <= c && c <= '9';

        /// <summary>Returns true IFF current character is between 'A' and 'Z, or Underbar or '$''.</summary>
        /// <remarks>These are the valid initial characters for an Identifier token.</remarks>
        public static bool IsAlpha       (this char c) => 'A' <= c && c <= 'Z' || c == '_' || c == '$';

        /// <summary>Returns true IFF current character is Alpha or Numeric.</summary>
        /// <remarks>These are the valid continuation characters for an Identifier token.</remarks>
        public static bool IsAlphanumeric(this char c) => c.IsAlpha() || c.IsNumeric();

        /// <summary>Returns true IFF current character is a LF, CR, TAB, FF, or SPACE.</summary>
        public static bool IsWhiteSpace  (this char c) => c == '\n' || c == '\r'
                                                       || c == '\t' || c == '\f' || c == ' ';

        public static bool IsWordOperator(this string text) => 
            LinksLexer.WordOperators.FirstOrDefault(s => s == text) != null;

        public static string Name(this IToken token) => token.Value.Name();

        public static void FastCopyToRange(this ITwoDimensionalLookup source, Excel.Range target) {
            var rowsCount = target.Rows.Count;
            var colsCount = target.Columns.Count;
            var data      = new object[rowsCount,colsCount];

            for(var row=0; row<rowsCount; row++)
                for(var col=0; col<colsCount; col++)
                    data[row,col] = source.Item(row, col);
            target.Value = data;
        }
    }
}
