////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.RibbonUtilities {
    /// <summary>.</summary>
    [CLSCompliant(false)]
    public class RibbonUtilitiesEntryPoint:IRibbonUtilities {
        /// <inheritdoc/>
        public ILinksAnalyzer NewLinksAnalyzer() => new LinksAnalyzer();

        public static VbaSourceExporter NewVbaSourceExporter() => new VbaSourceExporter(ExcelApp());

        private static Application ExcelApp() => new Application();
    }

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IRibbonUtilities {
        /// <summary>.</summary>
        ILinksAnalyzer NewLinksAnalyzer();
    }

    /// <summary>Static clas of ProgIds</summary>
    public static class ProgIds {
        /// <summary>ProgID for the Ribbon dispatcher.</summary>
        public const string RibbonUtilitiesProgId      = "PGSolutions.RibbonUtilities";
    }

    /// <summary>.</summary>
    /// <typeparam name="T"></typeparam>
    public class EventArgs<T>:EventArgs {
        public EventArgs(T value) : base() => Value = value;

        public T Value { get; }
    }
}
