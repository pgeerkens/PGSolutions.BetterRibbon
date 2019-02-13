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
        public RibbonToggleModel(Func<string, RibbonCheckBox> factory,
                IRibbonControlStrings strings, string imageMso, bool isEnabled, bool isVisible)
        : this(factory, strings, isEnabled, isVisible)
        => SetImageMso(imageMso);

        public RibbonToggleModel(Func<string, RibbonCheckBox> factory,
                IRibbonControlStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        : this(factory, strings, isEnabled, isVisible)
        => SetImageDisp(image);

        public RibbonToggleModel(Func<string, RibbonCheckBox> factory,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Factory = factory;
            Strings = strings;
            IsEnabled = isEnabled;
            IsEnabled = isVisible;
        }

        public IRibbonControlStrings Strings { get; }
        public bool   IsEnabled { get; set; } = true;
        public bool   IsVisible { get; set; } = true;

        #region IActivatable implementation
        private IRibbonToggle ViewModel { get; set; }

        private Func<string, RibbonCheckBox> Factory { get; }

        public IRibbonToggleModel Attach(string controlId) {
            var viewModel = Factory(controlId);
            if (viewModel != null) {
                viewModel.Attach(Getter).SetLanguageStrings(Strings);
                viewModel.Toggled += OnToggled;
            }
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        public void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                if (ViewModel is ISizeable sizeable) sizeable.SetSizeablel(this);
                if (ViewModel is IImageable imageable) imageable.SetImageable(this);

                ViewModel.Invalidate();
            }
        }
        #endregion

        #region IToggleable implementation
        public event ToggledEventHandler Toggled;

        private void OnToggled(object sender, bool isPressed)
        => Toggled?.Invoke(sender, IsPressed = isPressed);

        public bool IsPressed {
            get => _isPressed;
            set { _isPressed = value; ViewModel?.Invalidate(); }
        }
        bool _isPressed = false;

        public bool Getter() => IsPressed;
        #endregion

        #region ISizeable implementation
        public bool   IsLarge   { get; set; } = true;
        #endregion

        #region IImageable implementation
        public object Image     { get; set; } = "MacroSecurity";
        public bool   ShowImage { get; set; } = true;
        public bool   ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso) => Image = imageMso;
        #endregion
    }
}
