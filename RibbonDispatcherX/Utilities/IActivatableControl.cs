////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Utilities {
    public interface IActivatable {
        bool ShowWhenInactive { get; set; }
        void Detach();
        void Invalidate();
    }
    public interface IActivatableControl<TCtl> : IActivatable where TCtl : IRibbonCommon {
        TCtl Attach();
        new void Detach();
        new bool ShowWhenInactive { get; set; }
    }
    public interface IActivatableControl<TCtl, TSource> : IActivatable where TCtl:IRibbonCommon {
        TCtl Attach(Func<TSource> getter);
        new void Detach();
        new bool ShowWhenInactive { get; set; }
    }
}
