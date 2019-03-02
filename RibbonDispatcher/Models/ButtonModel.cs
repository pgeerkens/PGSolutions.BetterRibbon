////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon Button controls.</summary>
    [Description("The COM visible Model for Ribbon Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(IButtonModel))]
    [Guid(Guids.ButtonModel)]
    public class ButtonModel: ControlModel<IButtonSource,IButtonVM>, IButtonModel,
            IButtonSource {
        internal ButtonModel(Func<string, ButtonVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.Clicked += OnClicked; }
            return this;
        }

        #region IClickable implementation
        public event ClickedEventHandler Clicked;

        private void OnClicked(IRibbonControl control) => Clicked?.Invoke(control);
        #endregion

        public bool        IsLarge   { get; set; } = true;

        #region IImageable implementation
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IButtonModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion
    }
}
