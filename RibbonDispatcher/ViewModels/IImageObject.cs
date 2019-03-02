////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// </remarks>
    [Description("The main interface for VBA to access the Ribbon dispatcher.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IImageObject)]
    public interface IImageObject {
        string       ImageMso  { get; }
        IPictureDisp ImageDisp { get; }

        bool         IsMso     { get; }
    }
}
