using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>Interface exposed by an Excel workbook to the RibbonDispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonModel)]
    public interface IRibbonModel {
        IRibbonViewModel RibbonViewModel { get; }
        void ActivateTab();
    }
}