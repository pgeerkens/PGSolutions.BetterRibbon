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
    public sealed class RibbonButtonModel : RibbonControlModel<IRibbonButton>, IRibbonButtonModel, ISizeable, IImageable {
        public RibbonButtonModel(Func<string, RibbonButton> funcViewModel,
                IRibbonControlStrings strings, string imageMso, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageMso(imageMso);

        public RibbonButtonModel(Func<string, RibbonButton> funcViewModel,
                IRibbonControlStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageDisp(image);

        public RibbonButtonModel(Func<string,RibbonButton> funcViewModel, IRibbonControlStrings strings,
                bool isEnabled, bool isVisible) : base(strings, isEnabled, isVisible)
        => FuncViewModel = funcViewModel;

        public event ClickedEventHandler Clicked;

        public IRibbonButtonModel Attach(string controlId) {
            var viewModel = FuncViewModel(controlId);
            if (viewModel != null) {
                viewModel.Attach().SetLanguageStrings(Strings);
                viewModel.Clicked += OnClicked;
            }
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        private void OnClicked(object sender) => Clicked?.Invoke(sender);

        private Func<string, RibbonButton> FuncViewModel { get; }

        public bool   IsLarge   { get; set; } = true;
        public object Image     { get; set; } = "MacroSecurity";
        public bool   ShowImage { get; set; } = true;
        public bool   ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso)  => Image = imageMso;

        public override void Invalidate() {
            if (ViewModel != null) {
                if (ViewModel is ISizeable sizeable) sizeable.SetSizeablel(this);
                if (ViewModel is IImageable imageable) imageable.SetImageable(this);

                base.Invalidate();
            }
        }
    }
}
