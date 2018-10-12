////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    public interface IMain {
        /// <summary>TODO</summary>
        [Description("")]
        IRibbonFactory RibbonFactory { get; }

        /// <summary>TODO</summary>
        /// <param name="controlId"></param>
        /// <param name="strings"></param>
        /// <returns></returns>
        [Description("")]
        IRibbonButton AttachProxy(string controlId, IRibbonTextLanguageControl strings);

        /// <summary>TODO</summary>
        [Description("")]
        void DetachProxy(string controlId);

        /// <summary>TODO</summary>
        [Description("")]
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
