////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    [ComVisible(false)]
    public interface IActivatable {
        string Id           { get; }
        void   Detach();
        void   Invalidate();
        void   SetShowInactive(bool showInactive);
    }
    [ComVisible(false)]
    public interface IActivatable<TCtrl, TSource> : IActivatable
            where TCtrl : class,IRibbonCommon where TSource : IRibbonCommonSource {
        TCtrl Attach(TSource source);
        bool ShowInactive { get; }

        new string Id { get; }
        new void   Detach();
        new void   SetShowInactive(bool showInactive);
    }
}
