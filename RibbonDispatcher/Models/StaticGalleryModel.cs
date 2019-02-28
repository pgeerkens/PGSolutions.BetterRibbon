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
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IStaticGalleryModel))]
    [Guid(Guids.StaticGalleryModel)]
    public sealed class StaticGalleryModel : ControlModel<IStaticGallerySource,IStaticGalleryVM>, IStaticGalleryModel,
            IStaticGallerySource {
        internal StaticGalleryModel(Func<string, StaticGalleryVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IStaticGalleryModel Attach(string controlId) {
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
        private IReadOnlyList<StaticItemVM> Items => ViewModel.Items;

        public int Count => Items.Count;

        public int FindId(string id)
        => Items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();

        public IStaticItemVM this[int index] => Items[index];
        #endregion

        #region IGallerySize implementation
        /// <inheritdoc/>
        public int  ItemHeight { get; set; } = 15;

        /// <inheritdoc/>
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
