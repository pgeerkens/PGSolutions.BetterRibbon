////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IControlStrings;

    /// <summary>The COM visible Model for Ribbon Button controls.</summary>
    [Description("The COM visible Model for Ribbon Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISplitButtonModel))]
    [Guid(Guids.SplitButtonModel)]
    public class SplitButtonModel: ControlModel<ISplitButtonSource,ISplitButtonVM>,
            ISplitButtonModel, ISplitButtonSource {
        internal SplitButtonModel(Func<string, SplitButtonVM> funcViewModel, IStrings strings,
                ButtonModel button, MenuModel menu, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible) {
            Button = button;
            Menu   = menu;
        }

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public IButtonModel Button    { get; }
        public IMenuModel   Menu      { get; }

        public ISplitButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            Button.Attach(ViewModel.ButtonVM.Id);
            Menu.Attach(ViewModel.MenuVM.Id);
            ViewModel?.Invalidate();
            return this;
        }
    }
}
