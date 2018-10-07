﻿using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.Concrete;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
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
