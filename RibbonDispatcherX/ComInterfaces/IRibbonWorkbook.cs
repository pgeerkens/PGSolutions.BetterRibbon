using Microsoft.Office.Core;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>Interface exposed by an Excel workbook to the RibbonDispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonWorkbook)]
    public interface IRibbonWorkbook {
        /// <summary>The RibbonViewModel associated with this workbook.</summary>
        [DispId(1)]
        [Description("Returns the RibbonViewModel associated with this workbook.")]
        IRibbonViewModel ViewModel { get; }

        /// <summary>Initializes and returns a new RibbonModel for this {IRibbonUI}.</summary>
        [DispId(2)]
        [Description("Initializes and returns a new RibbonModel for this {IRibbonUI}.")]
        IRibbonModel InitializeRibbonModel(IRibbonUI ribbonUI);

        /// <summary>Returns the full path for this workbook.</summary>
        [DispId(3)]
        [Description("Returns the full path for this workbook.")]
        string Path { get; }
    }
}