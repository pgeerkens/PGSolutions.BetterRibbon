using Microsoft.Office.Core;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>Interface exposed by an Excel workbook to the RibbonDispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonWorkbook)]
    public interface IRibbonWorkbook {
        /// <summary>The RibbonViewModel associated with this workbook.</summary>
        [Description("Returns the RibbonViewModel associated with this workbook.")]
        IRibbonViewModel ViewModel { get; }

        /// <summary>Initializes and returns a new RibbonModel for this {IRibbonUI}.</summary>
        [Description("Initializes and returns a new RibbonModel for this {IRibbonUI}.")]
        IRibbonModel InitializeRibbonModel(IRibbonUI ribbonUI);

        /// <summary>Returns the full path for this workbook.</summary>
        [Description("Returns the full path for this workbook.")]
        string Path { get; }
    }
}