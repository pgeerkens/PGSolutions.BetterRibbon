////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon static DropDown controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable")]
    [SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IStaticDropDownModel))]
    [Guid(Guids.StaticDropDownModel)]
    public sealed class StaticDropDownModel : ControlModel<IStaticDropDownSource,IDropDownVM>, IStaticDropDownModel,
            IStaticDropDownSource, IEnumerable<ISelectableItemSource>, IEnumerable {
        internal StaticDropDownModel(Func<string, StaticDropDownVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IStaticDropDownModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        #region ISelectableList implementation
        public event SelectionMadeEventHandler SelectionMade;

        public int    SelectedIndex { get; set; }
        public string SelectedId    { get => Items[SelectedIndex].Id; set => SelectedIndex = FindId(value); }

        private void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex)
        => SelectionMade?.Invoke(control, selectedId, SelectedIndex = selectedIndex);
        #endregion

        #region IListable implementation
        private IList<ISelectableItemModel> Items { get; } = new List<ISelectableItemModel>();

        public int Count => Items.Count;

        public int FindId(string id)
        => Items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();

        //public IStaticItemVM this[int index] => Items[index];
        public ISelectableItemSource this[int index] => Items[index] as ISelectableItemSource;
        #endregion

        public IEnumerator<ISelectableItemSource> GetEnumerator() {
            foreach (var item in Items) yield return item as ISelectableItemSource;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
