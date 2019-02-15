////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    [ComDefaultInterface(typeof(IRibbonDropDownModel))]
    [Guid(Guids.RibbonDropDownModel)]
    public sealed class RibbonDropDownModel : RibbonControlModel<RibbonDropDown>, IRibbonDropDownModel,
            IRibbonDropDownSource {
        public RibbonDropDownModel(Func<string, RibbonDropDown> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event SelectedEventHandler SelectionMade;

        public int SelectedIndex  { get; set; }

        public IRibbonDropDownModel Attach(string controlId) {
            ViewModel = (FuncViewModel(controlId) as IActivatable<RibbonDropDown, IRibbonDropDownSource>)
                      ?.Attach(this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnSelectionMade(object sender, int selectedIndex)
        => SelectionMade?.Invoke(sender, SelectedIndex = selectedIndex);

        public IRibbonDropDownModel AddItem(ISelectableItem SelectableItem) {
            Items.Add(SelectableItem);
            ViewModel?.Invalidate();
            return this;
        }

        public IList<ISelectableItem>  Items { get; private set; } = new List<ISelectableItem>();
    }
}
