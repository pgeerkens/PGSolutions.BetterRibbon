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

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButtonModel))]
    [Guid(Guids.ButtonModel)]
    public class ButtonModel : RibbonControlModel<IRibbonButtonSource,ButtonVM>,
            IRibbonButtonModel, ISizeable, IImageable, IRibbonButtonSource {
        public ButtonModel(Func<string, ButtonVM> funcViewModel,
                IRibbonControlStrings strings, ImageObject image, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        => Image = image;

        public event EventHandler Clicked;

        public bool        IsLarge   { get; set; } = true;
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

         public IRibbonButtonModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Clicked += OnClicked;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnClicked(object sender, EventArgs e) => Clicked?.Invoke(sender,e);

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso)     => Image = imageMso;
    }
}
