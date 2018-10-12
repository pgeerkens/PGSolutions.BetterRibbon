////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComClasses;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    public interface IMain {
        IRibbonFactory RibbonFactory { get; }

        IRibbonButton AttachProxy(string controlId, IClickableRibbonButton proxy, IRibbonTextLanguageControl strings);

        void DetachProxy(string controlId);

        /// <inheritdoc/>
        void InvalidateControl(string ControlId);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringLoader
    {
        IRibbonTextLanguageControl GetStrings(string ControlId);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IImageLoader
    {
        IPictureDisp GetImage(string ControlId);
    }
}
//namespace PGSolutions.RibbonDispatcher.ComInterfaces {
//    internal static partial class DispIds {
//        public const int NewRibbonViewModel   = 1;
//        public const int SetRibbonUI          = 1 + NewRibbonViewModel;
//        public const int GetRibbonUI          = 1 + SetRibbonUI;
//    }
//}
