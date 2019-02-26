////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The COM visible Model for Ribbon EditBox controls.</summary>
    [Description("The COM visible Model for Ribbon EditBox controls.")]
    [CLSCompliant(true), ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvent))]
    [ComDefaultInterface(typeof(IToggleModel))]
    [Guid(Guids.ToggleModel)]
    public sealed class ToggleModel : ControlModel<IToggleSource, IToggleVM>,
            IToggleModel, IToggleSource {
        internal ToggleModel(Func<string, CheckBoxVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IToggleModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Toggled += OnToggled;
                ViewModel.Invalidate();
            }
            return this;
        }

        #region Toggleable implementation
        public event ToggledEventHandler Toggled;

        public bool        IsPressed { get; set; } = false;

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);
        #endregion

        #region ISizeable implementation
        public bool        IsLarge   { get; set; } = true;
        #endregion

        #region IImageable implementation
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso)     => Image = imageMso;
        #endregion
    }
}
