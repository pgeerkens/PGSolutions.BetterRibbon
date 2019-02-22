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
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The COM visible Model for Ribbon Drop Down controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectedEvents))]
    [ComDefaultInterface(typeof(IDropDownModel))]
    [Guid(Guids.DropDownModel)]
    public sealed class DropDownModel : RibbonControlModel<IDropDownSource,DropDownVM>,
            IDropDownModel, IDropDownSource, IEnumerable<ISelectableItem>, IEnumerable {
        public DropDownModel(Func<string, DropDownVM> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event SelectedEventHandler SelectionMade;

        public int SelectedIndex  { get; set; }

        public IDropDownModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnSelectionMade(object sender, int selectedIndex)
        => SelectionMade?.Invoke(sender, SelectedIndex = selectedIndex);

        public IDropDownModel AddSelectableModel(ISelectableItemModel selectableModel) {
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
