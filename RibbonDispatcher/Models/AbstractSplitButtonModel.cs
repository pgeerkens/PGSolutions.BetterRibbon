////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

    /// <summary>The COM visible Model for Ribbon Button controls.</summary>
    [Description("The COM visible Model for Ribbon Button controls.")]
    [CLSCompliant(true)]
    public abstract class AbstractSplitButtonModel<TSource,TControl>: ControlModel2<TSource,TControl>,
            ISplitButtonModel
        where TSource: IControlSource where TControl: class,ISplitButtonVM {
        protected AbstractSplitButtonModel(Func<string,IActivatable<TSource,TControl>> funcViewModel,
                IStrings2 strings, MenuModel menu)
        : base(funcViewModel, strings)
        => _menuModel = menu;

        public bool         IsLarge   { get; set; } = true;
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IMenuModel MenuModel => _menuModel; private MenuModel _menuModel { get; }

        protected void Attach(string controlId, TSource @this) {
            ViewModel = AttachToViewModel(controlId, @this);
            if (ViewModel != null) { MenuModel.Attach(ViewModel.MenuVM.ControlId); }
        }

        public override void Invalidate() { MenuModel.Invalidate(); base.Invalidate(); }

        public override void Detach() { MenuModel.Detach(); base.Detach(); }
    }
}
