////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using stdole;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon static Gallery controls.</summary>
    [Description("The COM visible Model for Ribbon static Gallery controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IGalleryModel))]
    [Guid(Guids.GalleryModel)]
    public sealed class GalleryModel : ControlModel<IGallerySource,IGalleryVM>, IGalleryModel,
            IGallerySource {
        internal GalleryModel(Func<string, GalleryVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        #region IDynamicListable implementation
        public IGalleryModel ClearList() { Items.Clear(); return this; }

        public IGalleryModel AddSelectableModel(ISelectableItemModel selectableModel) {
            Items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }
        #endregion

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

        public ISelectableItemSource this[int index] => Items[index] as ISelectableItemSource;
        #endregion

        #region IGallerySize implementation
        public int  ItemHeight { get; set; } = 15;
        public int  ItemWidth  { get; set; } = 15;
        #endregion

        #region IImageable implementation
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public void SetImageDisp(IPictureDisp image) => Image = new ImageObject(image);
        public void SetImageMso(string imageMso)     => Image = imageMso;
        #endregion
    }
}
