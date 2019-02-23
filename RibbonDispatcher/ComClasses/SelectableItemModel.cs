////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(ISelectableItemModel))]
    [Guid(Guids.SelectableItemModel)]
    public class SelectableItemModel : RibbonControlModel<ISelectableItemSource,SelectableItem>,
            ISelectableItemModel, ISizeable, IImageable, ISelectableItemSource, ISelectableItem {
        public SelectableItemModel(Func<string,SelectableItem> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible) { }

        public event ClickedEventHandler Clicked;

        public bool        IsLarge   { get => false; set { /* Not Supported - so ignore */ } } 
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public string Id        => ViewModel.Id;
        public string Label     => Strings.Label;
        public string ScreenTip => Strings.ScreenTip;
        public string SuperTip  => Strings.SuperTip;

        public ISelectableItemModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Clicked += OnClicked;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnClicked(IRibbonControl control) => Clicked?.Invoke(control);

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso) => Image = imageMso;
    }
}
