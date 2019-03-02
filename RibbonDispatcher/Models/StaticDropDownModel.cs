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
    /// <summary>The COM visible Model for Ribbon static DropDown controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IStaticDropDownModel))]
    [Guid(Guids.StaticDropDownModel)]
    public sealed class StaticDropDownModel : AbstractSelectableModel<IStaticDropDownSource,IDropDownVM>, IStaticDropDownModel,
            IStaticDropDownSource {
        internal StaticDropDownModel(Func<string, StaticDropDownVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IStaticDropDownModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.SelectionMade += OnSelectionMade; }
            return this;
        }
        #endregion

        #region IImageable implementation
        public IStaticDropDownModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion

        #region IListable implementation
        public override IReadOnlyList<IStaticItemVM> Items => ViewModel.Items;
        #endregion
    }
}
