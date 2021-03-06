////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

// This file is used by Code Analysis to maintain SuppressMessage 
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given 
// a specific target and scoped to a namespace, type, member, etc.
//
// To add a suppression to this file, right-click the message in the 
// Code Analysis results, point to "Suppress Message", and click 
// "In Suppression File".
// You do not need to add suppressions to this file manually.
using System.Diagnostics.CodeAnalysis;

[assembly: SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes",
        Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksLexer.#WordOperators")]

[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object,System.Object)", Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksParser.#ParseFormula(PGSolutions.LinksAnalysis.Interfaces.ISourceCellRef,System.String)" )]
[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)", Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksParser.#NewWorkbookNameRef(Microsoft.Office.Interop.Excel.Workbook,Microsoft.Office.Interop.Excel.Name)" )]
[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.Int32.ToString", Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksParser.#ExtendFromWorksheet(Microsoft.Office.Interop.Excel.Worksheet)" )]
[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object,System.Object)", Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksParser.#ExtendFromWorksheet(Microsoft.Office.Interop.Excel.Worksheet)" )]
[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object,System.Object)", Scope = "member", Target = "PGSolutions.LinksAnalysis.ExcelLinksExtensions.#WritePercentageStatus(Microsoft.Office.Interop.Excel.Worksheet,System.String,System.Int32)" )]
[assembly: SuppressMessage( "Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)", Scope = "member", Target = "PGSolutions.LinksAnalysis.LinksParser.#PGSolutions.LinksAnalysis.Interfaces.ITwoDimensionalLookup.Item(System.Int32,System.Int32)" )]
