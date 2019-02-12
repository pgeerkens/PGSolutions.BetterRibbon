////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonToggleModel)]
    public interface IRibbonToggleModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event ToggledEventHandler Toggled;

        bool   IsEnabled { get; set; }
        bool   IsVisible { get; set; }
        bool   IsLarge   { get; set; }
        object Image     { get; set; }
        bool   ShowImage { get; set; }
        bool   ShowLabel { get; set; }

        bool   IsPressed { get; set; }

        IRibbonToggleModel Attach(string controlId);

        void Invalidate();

        void SetImageDisp(IPictureDisp image);
        void SetImageMso(string imageMso);
    }
}
