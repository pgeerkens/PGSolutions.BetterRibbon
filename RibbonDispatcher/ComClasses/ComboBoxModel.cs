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
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The COM visible Model for Ribbon Drop Down controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent),typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IComboBoxModel))]
    [Guid(Guids.ComboBoxModel)]
    public sealed class ComboBoxModel: ControlModel<IComboBoxSource, IComboBoxVM>,
            IComboBoxModel, IComboBoxSource, IEnumerable<ISelectableItemModel>, IEnumerable {
        internal ComboBoxModel(Func<string, ComboBoxVM> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible) { }

        public event EditedEventHandler Edited;

        public event SelectionMadeEventHandler SelectionMade;

        public string Text          { get; set; } = "";

        public int    SelectedIndex { get; set; } = 0;

        public IComboBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited += OnEdited;
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control, text);

        private void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex)
        => SelectionMade?.Invoke(control, selectedId, SelectedIndex = selectedIndex);

        public IComboBoxModel AddSelectableModel(ISelectableItemModel selectableModel) {
            Items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }

        public ISelectableItemModel this[int index] => Items[index] as ISelectableItemModel;

        public int Count => Items.Count;

        private IList<ISelectableItemModel> Items { get; } = new List<ISelectableItemModel>();

        public IEnumerator<ISelectableItemModel> GetEnumerator() {
            foreach (var item in Items) yield return item as ISelectableItemModel;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
