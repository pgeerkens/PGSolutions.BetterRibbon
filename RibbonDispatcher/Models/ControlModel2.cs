////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>A variation on <see cref="ControlModel"/> exposing a <see cref="IStrings2"/> instead of an <see cref="IStrings"/>.</summary>
    /// <typeparam name="TSource"></typeparam>
    /// <typeparam name="TCtrl"></typeparam>
    public abstract class ControlModel2<TSource,TCtrl>: ControlModel<TSource,TCtrl>
            where TSource: IControlSource
            where TCtrl: class,IControlVM {
        protected ControlModel2(Func<string, IActivatable<TSource,TCtrl>> funcViewModel, IStrings2 strings)
        : base(funcViewModel, strings)=> Description = strings?.Description;

        /// <inheritdoc/>
        public string Description { get; set; }
    }
}
