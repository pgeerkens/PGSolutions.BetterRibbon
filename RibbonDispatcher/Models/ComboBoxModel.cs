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
using PGSolutions.RibbonDispatcher.ViewModels;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon Drop Down controls.</summary>
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent))]
    [ComDefaultInterface(typeof(IComboBoxModel))]
    [Guid(Guids.ComboBoxModel)]
    public sealed class ComboBoxModel: ControlModel<IComboBoxSource,IComboBoxVM>,
            IComboBoxModel, IComboBoxSource, IEnumerable<ISelectableItemSource>, IEnumerable {
        internal ComboBoxModel(Func<string, ComboBoxVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings) { }

        public event EditedEventHandler Edited;

        public string Text          { get; set; } = "";

        public IComboBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited += OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control, text);

        public IComboBoxModel AddSelectableModel(ISelectableItemModel selectableModel) {
            Items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }

        public ISelectableItemSource this[int index] => Items[index] as ISelectableItemSource;

        public int Count => Items.Count;

        private IList<ISelectableItemModel> Items { get; } = new List<ISelectableItemModel>();

        public IEnumerator<ISelectableItemSource> GetEnumerator() {
            foreach (var item in Items) yield return item as ISelectableItemSource;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
