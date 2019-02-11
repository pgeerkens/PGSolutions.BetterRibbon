////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

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
    [Guid(Guids.RibbonCheckboxModel)]
    public sealed class RibbonCheckboxModel : IRibbonToggleModel, IBooleanSource {
        public RibbonCheckboxModel(Func<string, RibbonCheckBox> factory) => Factory = factory;

        public event ToggledEventHandler Toggled;

        public IRibbonToggle ViewModel { get; private set; }

        public bool IsPressed {
            get => _isPressed;
            set { _isPressed = value; ViewModel.Invalidate(); }
        }
        bool _isPressed;

        public bool Getter() => IsPressed;

        public IRibbonToggleModel Attach(string controlId, IRibbonControlStrings strings) {
            var viewModel = Factory(controlId);
            viewModel.Attach(Getter).SetLanguageStrings(strings);
            viewModel.Toggled += OnToggled;
            ViewModel = viewModel;
            return this;
        }

        private void OnToggled(object sender, bool isPressed)
        => Toggled?.Invoke(sender, IsPressed = isPressed);

        private Func<string, RibbonCheckBox> Factory { get; }
    }
}
