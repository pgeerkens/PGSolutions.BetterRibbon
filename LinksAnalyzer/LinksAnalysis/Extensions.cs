////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Range = Microsoft.Office.Interop.Excel.Range;

    public static partial class Extensions {
        /// <summary>.</summary>
        /// <param name="range"></param>
        [CLSCompliant(false)]
        public static IReadOnlyList<string> GetNameList(this Range range) {
            if (range == null) return null;

            var list = new List<string>();

            foreach(Range cell in range) { list.Add(cell.Value2); }

            return list.AsReadOnly();
        }

        internal static IParseError ErrorHere(this ILinksLexer lexer, string source, string condition) =>
            new ParseError(lexer.CellRef, lexer.Formula, lexer.CharPosition, $"{source}: {condition}");

        /// <summary>Returns true IFF current character is between '0' and '9'.</summary>
        internal static bool IsNumeric     (this char c) => '0' <= c && c <= '9';

        /// <summary>Returns true IFF current character is between 'A' and 'Z, or Underbar or '$''.</summary>
        /// <remarks>These are the valid initial characters for an Identifier token.</remarks>
        internal static bool IsAlpha       (this char c) => 'A' <= c && c <= 'Z' || c == '_' || c == '$';

        /// <summary>Returns true IFF current character is Alpha or Numeric.</summary>
        /// <remarks>These are the valid continuation characters for an Identifier token.</remarks>
        internal static bool IsAlphanumeric(this char c) => c.IsAlpha() || c.IsNumeric();

        /// <summary>Returns true IFF current character is a LF, CR, TAB, FF, or SPACE.</summary>
        internal static bool IsWhiteSpace  (this char c) => c == '\n' || c == '\r'
                                                         || c == '\t' || c == '\f' || c == ' ';

        internal static bool IsWordOperator(this string text) => 
            LinksLexer.WordOperators.FirstOrDefault(s => s == text) != null;

        internal static string Name(this IToken token) => token.Value.Name();

        [SuppressMessage( "Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", MessageId = "Body" )]
        [CLSCompliant(false)]
        public static void FastCopyToRange(this IReadOnlyList<ICellRef> source, Range target) {
            if (source==null) throw new ArgumentNullException(nameof(source));
            if (target==null) throw new ArgumentNullException(nameof(target));

            var rowsCount = target.Rows.Count;
            var colsCount = target.Columns.Count;
            var data      = new object[rowsCount,colsCount];

            for(var row=0; row<rowsCount; row++) {
                for (var col=0; col<colsCount; col++) {
                   data[row,col] = source.Item(row, col); 
                }
            }
            target.Value = data;
        }
    }
}
