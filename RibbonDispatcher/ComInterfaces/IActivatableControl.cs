////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    [ComVisible(false)]
    public interface IActivatable {
        string Id { get; }
        bool ShowActiveOnly { get; set; }
        void Detach();
        void Invalidate();
    }
    [ComVisible(false)]
    public interface IActivatableControl<TCtl> : IActivatable where TCtl : IRibbonCommon {
        new string Id { get; }
        TCtl Attach();
        new void Detach();
        new bool ShowActiveOnly { get; set; }
    }
    [ComVisible(false)]
    public interface IActivatableControl<TCtl, TSource> : IActivatable where TCtl:IRibbonCommon {
        new string Id { get; }
        TCtl Attach(Func<TSource> getter);
        new void Detach();
        new bool ShowActiveOnly { get; set; }
    }
}
