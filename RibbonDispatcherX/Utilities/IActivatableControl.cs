////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.Utilities {
    public interface IActivatableControl<TCtl> where TCtl : IRibbonCommon {
        TCtl Attach();
        void Detach();
    }
    public interface IActivatableControl<TCtl, TSource> where TCtl:IRibbonCommon {
        TCtl Attach(Func<TSource> getter);
        void Detach();
    }
}
