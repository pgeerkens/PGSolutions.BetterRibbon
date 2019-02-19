////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
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
    public sealed class RibbonDropDownModel : RibbonControlModel<IRibbonDropDownSource,RibbonDropDown>,
            IRibbonDropDownModel, IRibbonDropDownSource, IEnumerable<ISelectableItem>, IEnumerable {
        public RibbonDropDownModel(Func<string, RibbonDropDown> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event SelectedEventHandler SelectionMade;

        public int SelectedIndex  { get; set; }

        public IRibbonDropDownModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnSelectionMade(object sender, int selectedIndex)
        => SelectionMade?.Invoke(sender, SelectedIndex = selectedIndex);

        public IRibbonDropDownModel AddSelectableModel(ISelectableItemModel selectableModel) {
            Items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }

        public ISelectableItem this[int index] => Items[index] as ISelectableItem;

        public int Count => Items.Count;

        private IList<ISelectableItemModel> Items { get; } = new List<ISelectableItemModel>();

        public IEnumerator<ISelectableItem> GetEnumerator() {
            foreach (var item in Items) yield return item as ISelectableItem;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
