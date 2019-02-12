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
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButtonModel))]
    [Guid(Guids.RibbonButtonModel)]
    public sealed class RibbonButtonModel : IRibbonButtonModel, ISizeable, IImageable {
        public RibbonButtonModel(Func<string, RibbonButton> factory,
                IRibbonControlStrings strings, string imageMso, bool isEnabled, bool isVisible)
        : this(factory, strings, isEnabled, isVisible)
        => SetImageMso(imageMso);

        public RibbonButtonModel(Func<string, RibbonButton> factory,
                IRibbonControlStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        : this(factory, strings, isEnabled, isVisible)
        => SetImageDisp(image);

        public RibbonButtonModel(Func<string,RibbonButton> factory, IRibbonControlStrings strings,
                bool isEnabled, bool isVisible) {
            Factory = factory;
            Strings = strings;
            IsEnabled = isEnabled;
            IsEnabled = isVisible;
        }

        public event ClickedEventHandler Clicked;

        IRibbonButton ViewModel { get; set; }

        public IRibbonButtonModel Attach(string controlId) {
            var viewModel = Factory(controlId);
            if (viewModel != null) {
                viewModel.Attach().SetLanguageStrings(Strings);
                viewModel.Clicked += OnClicked;
            }
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        private void OnClicked(object sender) => Clicked(sender);

        private Func<string, RibbonButton> Factory { get; }

        public IRibbonControlStrings Strings { get; }
        public bool   IsEnabled { get; set; } = true;
        public bool   IsVisible { get; set; } = true;
        public bool   IsLarge   { get; set; } = true;
        public object Image     { get; set; } = "MacroSecurity";
        public bool   ShowImage { get; set; } = true;
        public bool   ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso)  => Image = imageMso;

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
