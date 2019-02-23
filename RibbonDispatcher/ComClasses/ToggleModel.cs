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
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IToggleModel))]
    [Guid(Guids.ToggleModel)]
    public sealed class ToggleModel : RibbonControlModel<IRibbonToggleSource,CheckBoxVM>,
            IToggleModel, ISizeable, IImageable, IRibbonToggleSource {
        public ToggleModel(Func<string, CheckBoxVM> funcViewModel,
                IControlStrings strings, ImageObject image, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        => Image = image;

        public event ToggledEventHandler Toggled;

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public bool        IsPressed { get; set; } = false;

        public IToggleModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Toggled += OnToggled;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control, IsPressed = isPressed);

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso)     => Image = imageMso;
    }
}
