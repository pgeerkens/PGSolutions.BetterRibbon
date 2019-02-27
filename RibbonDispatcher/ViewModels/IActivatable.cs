////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The contract specifying how ViewModel classes are attached to by a <typeparamref name="TSource"/> Model."/></summary>
    /// <typeparam name="TSource">The contract type required of Model classes desiring to attach.</typeparam>
    /// <typeparam name="TControl">The class type of the ViewModel class implementing this interface.</typeparam>
    [ComVisible(false)]
    public interface IActivatable<TSource, TControl>: IControlVM
            where TControl: IControlVM
            where TSource: IControlSource {
        bool     ShowInactive { get; }

        TControl Attach(TSource source);
    }
}
