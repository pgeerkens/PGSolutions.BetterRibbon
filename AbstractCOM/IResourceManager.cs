using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IResourceManager)]
    public interface IResourceManager {
        /// <summary>TODO</summary>
        [DispId(1)]
        [Description("Returns the resource string of the given name.")]
        string GetCurrentUIString(string name);
        /// <summary>TODO</summary>
        [DispId(2)]
        [Description("Returns the image associated with the supplied name.")]
        object LoadImage(string name);
    }
    /// <summary>Interface exposed by an Excel workbook to the RibbonDispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonWorkbook)]
    public interface IRibbonWorkbook {
        /// <summary>The RibbonViewModel associated with this workbook.</summary>
        [DispId(1)]
        [Description("Returns the RibbonViewModel associated with this workbook.")]
        RibbonViewModel ViewModel { get; }
    }
}
