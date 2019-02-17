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
    public class RibbonButtonModel : RibbonControlModel<RibbonButton>, IRibbonButtonModel,
            ISizeable, IImageable, IRibbonButtonSource {
        public RibbonButtonModel(Func<string, RibbonButton> funcViewModel,
                IRibbonControlStrings strings, ImageObject image, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        => Image = image;

        public event ClickedEventHandler Clicked;

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public IRibbonButtonModel Attach(string controlId) => Attach(FuncViewModel(controlId));

        internal IRibbonButtonModel Attach(RibbonButton viewModel) {
            (viewModel as IActivatable<RibbonButton, IRibbonButtonSource>)?.Attach(this);
            if (viewModel != null) {
                ViewModel = viewModel;
                ViewModel.Clicked += OnClicked;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnClicked(object sender) => Clicked?.Invoke(sender);

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso)     => Image = imageMso;
    }
}
namespace PGSolutions.RibbonDispatcher.ComInterfaces {
}
