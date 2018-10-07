using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher2013.AbstractCOM {
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
    }
}
