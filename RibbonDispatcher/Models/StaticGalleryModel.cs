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
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IStaticGalleryModel))]
    [Guid(Guids.StaticGalleryModel)]
    public sealed class StaticGalleryModel : AbstractSelectableModel<IStaticGallerySource,IStaticGalleryVM>, IStaticGalleryModel,
            IStaticGallerySource {
        internal StaticGalleryModel(Func<string, StaticGalleryVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IStaticGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }
        #endregion

        #region IListable implementation
        public override IReadOnlyList<IStaticItemVM> Items => ViewModel.Items;
        #endregion

        #region IGallerySize implementation
        public int  ItemHeight { get; set; } = 15;
        public int  ItemWidth  { get; set; } = 15;
        #endregion
    }
}
