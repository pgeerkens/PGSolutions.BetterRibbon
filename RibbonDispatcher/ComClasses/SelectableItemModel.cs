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
    [Guid(Guids.SelectableItemModel)]
    public class SelectableItemModel : RibbonControlModel<SelectableItem>, ISelectableItem,
             IRibbonButtonModel, ISizeable, IImageable, IRibbonButtonSource {
        public SelectableItemModel(Func<string, SelectableItem> funcViewModel,
                IRibbonControlStrings strings, string imageMso, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageMso(imageMso);

        public SelectableItemModel(Func<string, SelectableItem> funcViewModel,
                IRibbonControlStrings strings, IPictureDisp image, bool isEnabled, bool isVisible)
        : this(funcViewModel, strings, isEnabled, isVisible)
        => SetImageDisp(image);

        public SelectableItemModel(Func<string, SelectableItem> funcViewModel, IRibbonControlStrings strings,
                bool isEnabled, bool isVisible) : base(strings, isEnabled, isVisible)
        => FuncViewModel = funcViewModel;

        public event ClickedEventHandler Clicked;

        public string Id        => ViewModel.Id;
        public string Label     => Strings.Label;
        public string ScreenTip => Strings.ScreenTip;
        public string SuperTip  => Strings.SuperTip;

        public bool IsLarge { get; set; } = true;
        public object Image { get; set; } = "MacroSecurity";
        public bool ShowImage { get; set; } = true;
        public bool ShowLabel { get; set; } = true;

        private Func<string, SelectableItem> FuncViewModel { get; }

        public IRibbonButtonModel Attach(string controlId) {
            ViewModel = (FuncViewModel(controlId) as IActivatable<SelectableItem, IRibbonButtonSource>)
                      ?.Attach(this);
            if (ViewModel != null) {
            //    ViewModel.Clicked += OnClicked;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnClicked(object sender) => Clicked?.Invoke(sender);

        public void SetImageDisp(IPictureDisp image) => Image = image;
        public void SetImageMso(string imageMso) => Image = imageMso;
    }
}
