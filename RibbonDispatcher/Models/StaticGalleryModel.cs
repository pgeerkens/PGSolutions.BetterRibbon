﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

   /// <summary>The COM visible Model for Ribbon static Gallery controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IStaticGalleryModel))]
    [Guid(Guids.StaticGalleryModel)]
    public sealed class StaticGalleryModel : AbstractSelectableModel2<IStaticGallerySource,IStaticGalleryVM>, IStaticGalleryModel,
            IStaticGallerySource {
        internal StaticGalleryModel(Func<string, StaticGalleryVM> funcViewModel, IStrings2 strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IStaticGalleryModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.SelectionMade += OnSelectionMade; }
            return this;
        }
        #endregion

        #region ISizeable implementation
        public bool        IsLarge   { get; set; } = true;
        #endregion

        #region IImageable implementation
        public IStaticGalleryModel SetImage(IImageObject image) { Image = image; return this; }
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
