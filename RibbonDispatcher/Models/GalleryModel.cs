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
    /// <summary>The COM visible Model for Ribbon Drop Down controls.</summary>
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable")]
    [SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible")]
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IGalleryModel))]
    [Guid(Guids.GalleryModel)]
    public sealed class GalleryModel : ControlModel<IGallerySource,IGalleryVM>,
            IGalleryModel, IGallerySource, IEnumerable<ISelectableItemSource>, IEnumerable {
        internal GalleryModel(Func<string, GalleryVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public event SelectionMadeEventHandler SelectionMade;

        public int    SelectedIndex { get; set; }
        public string SelectedId    {
            get => Items[SelectedIndex].Id;
            set => SelectedIndex = Items.Where((item,i) => item.Id == value).Select((a,b)=>b).FirstOrDefault();
        }

        public IGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex)
        => SelectionMade?.Invoke(control, SelectedId = selectedId, SelectedIndex = selectedIndex);

        public IGalleryModel AddSelectableModel(ISelectableItemModel selectableModel) {
            Items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }

        /// <inheritdoc/>
        public int ItemHeight { get; set; }

        /// <inheritdoc/>
        public int ItemWidth  { get; set; }

        public ISelectableItemSource this[int index] => Items[index] as ISelectableItemSource;

        public int Count => Items.Count;

        private IList<ISelectableItemModel> Items { get; } = new List<ISelectableItemModel>();

        public IEnumerator<ISelectableItemSource> GetEnumerator() {
            foreach (var item in Items) yield return item as ISelectableItemSource;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
