////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggleModel))]
    [Guid(Guids.RibbonToggleModel)]
    public sealed class RibbonToggleModel : IRibbonToggleModel, IBooleanSource, ISizeable, IImageable {
        public RibbonToggleModel(Func<string, RibbonCheckBox> factory) => Factory = factory;

        public event ToggledEventHandler Toggled;

        IRibbonToggle ViewModel { get; set; }

        public bool IsPressed {
            get => _isPressed;
            set { _isPressed = value; ViewModel?.Invalidate(); }
        } bool _isPressed = false;

        public bool Getter() => IsPressed;

        public IRibbonToggleModel Attach(string controlId, IRibbonControlStrings strings) {
            var viewModel = Factory(controlId);
            if (viewModel != null) {
                viewModel.Attach(Getter).SetLanguageStrings(strings);
                viewModel.Toggled += OnToggled;
            }
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        private void OnToggled(object sender, bool isPressed)
        => Toggled?.Invoke(sender, IsPressed = isPressed);

        private Func<string, RibbonCheckBox> Factory { get; }

        public bool   IsEnabled { get; set; } = true;
        public bool   IsVisible { get; set; } = true;
        public bool   IsLarge   { get; set; } = true;
        public object Image     { get; set; } = "MacroSecurity";
        public bool   ShowImage { get; set; } = true;
        public bool   ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso) => Image = imageMso;

        public void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                if (ViewModel is ISizeable sizeable)   sizeable.SetSizeablel(this);
                if (ViewModel is IImageable imageable) imageable.SetImageable(this);

                ViewModel.Invalidate();
            }
        }
    }
}
