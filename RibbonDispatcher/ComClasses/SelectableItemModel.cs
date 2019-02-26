////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(ISelectableItemModel))]
    [Guid(Guids.SelectableItemModel)]
    public class SelectableItemModel : ControlModel<ISelectableItemSource,ISelectableItemVM>,
            ISelectableItemModel, ISelectableItemSource, ISelectableItemVM {
        internal SelectableItemModel(IControlStrings strings)
        : base(null, strings) { }

        public bool        IsLarge   { get => false; set { /* Not Supported - so ignore */ } } 
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public string Id        { get; set; } = null;
        public string Label     => Strings.Label;
        public string ScreenTip => Strings.ScreenTip;
        public string SuperTip  => Strings.SuperTip;
        public string KeyTip    => Strings.KeyTip;

        public new ISelectableItemVM ViewModel => this;

        public ISelectableItemModel Attach(string controlId) {
            Id = controlId;
            Invalidate();
            return this;
        }

        public override void Detach() { Id = null; base.Detach(); }

        public override void Invalidate() { }

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso) => Image = imageMso;
    }
}
