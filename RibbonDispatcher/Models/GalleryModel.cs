////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

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
    public sealed class GalleryModel : AbstractSelectableModel<IGallerySource,IGalleryVM>, IGalleryModel,
            IGallerySource {
        internal GalleryModel(Func<string, GalleryVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }
        #endregion

        #region IListable implementation
        public override IReadOnlyList<IStaticItemVM> Items => _items.AsReadOnly();
        private List<IStaticItemVM> _items = new List<IStaticItemVM>();
        #endregion

        #region IDynamicListable implementation
        public IGalleryModel ClearList() { _items.Clear(); return this; }

        public IGalleryModel AddSelectableModel(IStaticItemVM selectableModel) {
            _items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }
        #endregion

        #region IGallerySize implementation
        public int  ItemHeight { get; set; } = 15;
        public int  ItemWidth  { get; set; } = 15;
        #endregion
    }
}
