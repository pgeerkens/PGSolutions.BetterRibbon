////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

    /// <summary>A variation on <see cref="AbstractSelectableModel"/> exposing Description text.</summary>
    public abstract class AbstractSelectableModel2<TSource,TCtrl> : AbstractSelectableModel<TSource,TCtrl>, IControlSource
            where TSource: IControlSource
            where TCtrl: class,IControlVM  {
        internal AbstractSelectableModel2(Func<string, IActivatable<TSource, TCtrl>> funcViewModel, IStrings2 strings)
        : base(funcViewModel,strings) => Description = strings?.Description;

        public string Description { set; get; }
    }
}
