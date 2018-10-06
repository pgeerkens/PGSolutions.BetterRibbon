////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.AbstractCOM;

namespace PGSolutions.RibbonDispatcher {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    public interface IMain {
        /// <summary>Returns a new instance of {RibbonViewModel} for the supplied {IRibbonUI} and {IResourceManager}.</summary>
        [DispId(1)]
        [Description("Returns a new instance of {RibbonViewModel} for the supplied {IRibbonUI} and {IResourceManager}.")]
        IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, IResourceManager resourceManager);

        /// <summary>TODO</summary>
        IRibbonUI SetRibbonUI(IRibbonUI ribbonUI, string workbookPath);

        /// <summary>TODO</summary>
        IRibbonUI GetRibbonUI(string WorkbookPath);
    }
}
