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
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(ISelectableItemModel))]
    [Guid(Guids.SelectableItemModel)]
    public class SelectableItemModel : ControlModel<ISelectableItemSource,ISelectableItemVM>,
            ISelectableItemModel, ISelectableItemSource, ISelectableItemVM {
        internal SelectableItemModel(
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(null, strings, isEnabled, isVisible) { }

        public bool        IsLarge   { get => false; set { /* Not Supported - so ignore */ } } 
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public string Id        { get; set; } = null;
        public string Label     => Strings.Label;
        public string ScreenTip => Strings.ScreenTip;
        public string SuperTip  => Strings.SuperTip;

        public new ISelectableItemVM ViewModel => this;

        string IControlVM.Description => Strings.Description;

        string IControlVM.KeyTip      => Strings.KeyTip;

        string IControlVM.Label       => Strings.Label;

        string IControlVM.ScreenTip   => Strings.ScreenTip;

        string IControlVM.SuperTip    => Strings.SuperTip;

        public ISelectableItemModel Attach(string controlId) {
            Id = controlId;
            Invalidate();
            return this;
        }

        public void Detach() { Id = null; Invalidate(); }

        public override void Invalidate() { }

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso) => Image = imageMso;
    }
}
