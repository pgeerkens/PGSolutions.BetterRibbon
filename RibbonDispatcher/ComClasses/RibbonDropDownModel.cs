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
    public sealed class RibbonDropDownModel : IRibbonDropDownModel, IIntegerSource {
        public RibbonDropDownModel(Func<string, RibbonDropDown> factory) => Factory = factory;

        public event SelectedEventHandler SelectionMade;

        private IList<ISelectableItem>  _items  = new List<ISelectableItem>();

        IRibbonDropDown ViewModel { get; set; }

        public int SelectedIndex  { get; set; }

        public int Getter() => SelectedIndex;

        public IRibbonDropDownModel AddItem(ISelectableItem SelectableItem) {
            _items.Add(SelectableItem);
            return this;
        }

        public IRibbonDropDownModel Attach(string controlId, IRibbonControlStrings strings) {
            var viewModel = Factory(controlId);
            if (viewModel != null) {
                viewModel.Attach(Getter).SetLanguageStrings(strings);
                viewModel.SelectionMade += OnSelectionMade;
                foreach (var item in _items) viewModel.AddItem(item);
            }
            ViewModel = viewModel;
            Invalidate();
            return this;
        }

        private void OnSelectionMade(object sender, int selectedIndex)
        => SelectionMade?.Invoke(sender, SelectedIndex = selectedIndex);

        private Func<string, RibbonDropDown> Factory { get; }

        public bool IsEnabled { get; set; } = true;
        public bool IsVisible { get; set; } = true;

        public void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                //if (ViewModel is ISizeable sizeable)   sizeable.SetSizeablel(this);
                //if (ViewModel is IImageable imageable) imageable.SetImageable(this);

                ViewModel.Invalidate();
            }
        }
    }
}
