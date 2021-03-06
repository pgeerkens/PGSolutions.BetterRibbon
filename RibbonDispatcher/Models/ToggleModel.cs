﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

    /// <summary>The COM visible Model for Ribbon ToggleButton and CHeckBox controls.</summary>
    [Description("The COM visible Model for Ribbon ToggleButton and CHeckBox controls.")]
    [CLSCompliant(true), ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvent))]
    [ComDefaultInterface(typeof(IToggleModel))]
    [Guid(Guids.ToggleModel)]
    public sealed class ToggleModel : ControlModel2<IToggleSource, IToggleVM>,
            IToggleModel, IToggleSource {
        internal ToggleModel(Func<string, CheckBoxVM> funcViewModel, IStrings2 strings)
        : base(funcViewModel, strings) { }

        public IToggleModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.Toggled += OnToggled; }
            return this;
        }

        #region Toggleable implementation
        public event ToggledEventHandler Toggled;

        public bool        IsPressed { get; set; } = false;

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);
        #endregion

        #region ISizeable implementation
        public bool        IsLarge   { get; set; } = false;
        #endregion

        #region IImageable implementation
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IToggleModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion
    }
}
