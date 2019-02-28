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
    /// <summary>The COM visible Model for Ribbon Drop Down controls.</summary>
    [Description("The COM visible Model for Ribbon Drop Down controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvent))]
    [ComDefaultInterface(typeof(IDropDownModel))]
    [Guid(Guids.DropDownModel)]
    public sealed class DropDownModel : AbstractSelectableModel<IDropDownSource,IDropDownVM>, IDropDownModel,
            IDropDownSource {
        internal DropDownModel(Func<string, DropDownVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IDropDownModel Attach(string controlId) {
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
        #endregion

        #region IDynamicListable implementation
        private List<IStaticItemVM> _items = new List<IStaticItemVM>();

        public IDropDownModel ClearList() { _items.Clear(); return this; }

        public IDropDownModel AddSelectableModel(IStaticItemVM selectableModel) {
            _items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }
        #endregion
    }
}
