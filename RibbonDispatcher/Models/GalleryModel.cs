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
    using IStrings = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>The COM visible Model for Ribbon static Gallery controls.</summary>
    [Description("The COM visible Model for Ribbon static Gallery controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IGalleryModel))]
    [Guid(Guids.GalleryModel)]
    public sealed class GalleryModel : AbstractSelectableModel2<IGallerySource,IGalleryVM>, IGalleryModel,
            IGallerySource {
        internal GalleryModel(Func<string, GalleryVM> funcViewModel, IStrings2 strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.SelectionMade += OnSelectionMade; }
            return this;
        }
        #endregion

        public bool        IsLarge   { get; set; } = true;

        #region IImageable implementation
        public IGalleryModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion

        #region IListable implementation
        public override IReadOnlyList<IStaticItemVM> Items => _items.AsReadOnly();
        private List<IStaticItemVM> _items = new List<IStaticItemVM>();
        #endregion

        #region IDynamicListable implementation
        public IGalleryModel ClearList() { _items.Clear(); return this; }

        public IGalleryModel AddSelectableModel(IStaticItemVM selectableModel) {
            _items.Add(selectableModel);
            return this;
        }
        #endregion

        #region IGallerySize implementation
        public int  ItemHeight { get; set; } = 15;
        public int  ItemWidth  { get; set; } = 15;
        #endregion
    }
}
