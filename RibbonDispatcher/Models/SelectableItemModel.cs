﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary></summary>
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(ISelectableItemModel))]
    [Guid(Guids.SelectableItemModel)]
    public class SelectableItemModel : ControlModel<ISelectableItemSource,IStaticItemVM>,
            ISelectableItemModel, ISelectableItemSource, IStaticItemVM {
        internal SelectableItemModel(IControlStrings strings)
        : base(null, strings) { }

        public bool         IsLarge   { get => false; set { /* Not Supported - so ignore */ } } 
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public string ControlId { get; set; } = null;

        public new IStaticItemVM ViewModel => this;

        public ISelectableItemModel Attach(string controlId) {
            ControlId = controlId;
            return this;
        }

        public override void Detach() { ControlId = null; base.Detach(); }

        public override void Invalidate() { }

        public void OnPurged(IContainerControl sender) { }  // Own VM!

        #region IImageable implementation
        public ISelectableItemModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion
    }
}
