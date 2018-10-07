////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    public interface IMain {
        /// <summary>Returns a new instance of {RibbonViewModel} for the supplied {IRibbonUI} and {IResourceManager}.</summary>
        [DispId(DispIds.NewRibbonViewModel)]
        [Description("Returns a new instance of {RibbonViewModel} for the supplied {IRibbonUI} and {IResourceManager}.")]
        IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI);

        /// <summary>Adds the supplied {IRibbonUI} to an in-memory cache using supplied workbookPath as a key.</summary>
        [DispId(DispIds.SetRibbonUI)]
        [Description("Adds the supplied {IRibbonUI} to an in-memory cache using supplied workbookPath as a key.")]
        IRibbonUI SetRibbonUI(IRibbonUI ribbonUI, string workbookPath);

        /// <summary>Retrieves a {IRibbonUI} keyed by the supplied workbookPath from the in-memory cache.</summary>
        [DispId(DispIds.GetRibbonUI)]
        [Description("Retrieves a {IRibbonUI} keyed by the supplied workbookPath from the in-memory cache.")]
        IRibbonUI GetRibbonUI(string WorkbookPath);
    }

    internal static partial class DispIds {
        public const int NewRibbonViewModel   = 1;
        public const int SetRibbonUI          = 1 + NewRibbonViewModel;
        public const int GetRibbonUI          = 1 + SetRibbonUI;
    }
}
