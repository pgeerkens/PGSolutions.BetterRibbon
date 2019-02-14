﻿////////////////////////////////////////////////////////////////////////////////////////////////////
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
    public sealed class RibbonDropDownModel : RibbonControlModel<IRibbonDropDown>, IRibbonDropDownModel, IIntegerSource {
        public RibbonDropDownModel(Func<string, RibbonDropDown> factory, IRibbonControlStrings strings,
                bool isEnabled, bool isVisible)
        :base(strings, isEnabled, isVisible)
        => Factory = factory;

        public event SelectedEventHandler SelectionMade;

        private IList<ISelectableItem>  _items  = new List<ISelectableItem>();

        public int SelectedIndex  { get; set; }

        public int Getter() => SelectedIndex;

        public IRibbonDropDownModel AddItem(ISelectableItem SelectableItem) {
            _items.Add(SelectableItem);
            return this;
        }

        public IRibbonDropDownModel Attach(string controlId) {
            var viewModel = Factory(controlId);
            if (viewModel != null) {
                viewModel.Attach(Getter).SetLanguageStrings(Strings);
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
    }
}
