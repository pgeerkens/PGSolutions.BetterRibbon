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
    public sealed class RibbonToggleModel : RibbonControlModel<RibbonCheckBox>, IRibbonToggleModel,
                ISizeable, IImageable, IRibbonToggleSource {
        public RibbonToggleModel(Func<string, RibbonCheckBox> funcViewModel,
                IRibbonControlStrings strings, string imageMso, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageMso(imageMso);

        public RibbonToggleModel(Func<string, RibbonCheckBox> funcViewModel,
                IRibbonControlStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageDisp(image);

        public RibbonToggleModel(Func<string, RibbonCheckBox> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible) { }

        public event ToggledEventHandler Toggled;

        public bool   IsLarge   { get; set; } = true;
        public object Image     { get; set; } = "MacroSecurity";
        public bool   ShowImage { get; set; } = true;
        public bool   ShowLabel { get; set; } = true;

        public bool   IsPressed { get; set; } = false;

        public IRibbonToggleModel Attach(string controlId) {
            ViewModel = (FuncViewModel(controlId) as IActivatable<RibbonCheckBox, IRibbonToggleSource>)
                      ?.Attach(this);
            if (ViewModel != null) {
                ViewModel.Toggled += OnToggled;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnToggled(object sender, bool isPressed) => Toggled?.Invoke(sender, IsPressed = isPressed);

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso)     => Image = imageMso;
    }
}
