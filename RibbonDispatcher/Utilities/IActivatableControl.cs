////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Utilities {
    [ComVisible(false)]
    public interface IActivatable {
        bool ShowWhenInactive { get; set; }
        void Detach();
        void Invalidate();
    }
    [ComVisible(false)]
    public interface IActivatableControl<TCtl> : IActivatable where TCtl : IRibbonCommon {
        TCtl Attach();
        new void Detach();
        new bool ShowWhenInactive { get; set; }
    }
    [ComVisible(false)]
    public interface IActivatableControl<TCtl, TSource> : IActivatable where TCtl:IRibbonCommon {
        TCtl Attach(Func<TSource> getter);
        new void Detach();
        new bool ShowWhenInactive { get; set; }
    }
}
