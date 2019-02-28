﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;

    /// <summary>The COM visible Model for Ribbon Button controls.</summary>
    [Description("The COM visible Model for Ribbon Button controls.")]
    [CLSCompliant(true)]
    public abstract class AbstractSplitButtonModel<TSource,TControl>: ControlModel<TSource,TControl>,
            ISplitButtonModel
        where TSource: IControlSource where TControl: ISplitButtonVM {
        protected AbstractSplitButtonModel(Func<string,IActivatable<TSource,TControl>> funcViewModel, IStrings strings,
                MenuModel menu)
        : base(funcViewModel, strings)
        => Menu   = menu;

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public MenuModel   Menu      { get; }

        public override void Detach() { Menu.Detach(); base.Detach(); }
    }
}
