////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The contract specifying that a ViewModel implementation can be attached to bya Model.</summary>
    [ComVisible(false)]
    public interface IActivatable {
        string Id { get; }

        void   Detach();
        void   Invalidate();
        void   SetShowInactive(bool showInactive);
    }

    /// <summary>The contract specifying how ViewModel classes are attached to by a <typeparamref name="TSource"/> Model."/></summary>
    /// <typeparam name="TSource">The contract type required of Model classes desiring to attach.</typeparam>
    /// <typeparam name="TControl">The class type of the ViewModel class implementing this interface.</typeparam>
    [ComVisible(false)]
    public interface IActivatable<TSource, TControl>: IActivatable
            where TControl: IControlVM
            where TSource: IRibbonCommonSource {
        bool     ShowInactive { get; }

        TControl Attach(TSource source);
    }
}
