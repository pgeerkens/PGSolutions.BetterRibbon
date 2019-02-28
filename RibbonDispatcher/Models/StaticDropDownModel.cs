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
            IStaticDropDownSource {
        internal StaticDropDownModel(Func<string, StaticDropDownVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IStaticDropDownModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.SelectionMade += OnSelectionMade;
                ViewModel.Invalidate();
            }
            return this;
        }

        public override void Detach() { SelectionMade = null;  base.Detach(); }
        #endregion

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items => ViewModel.Items;

        public int FindId(string id)
        => Items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();
        #endregion

        #region ISelectableList implementation
        public event SelectionMadeEventHandler SelectionMade;

        public int    SelectedIndex { get; set; }
        public string SelectedId    { get => Items[SelectedIndex].Id; set => SelectedIndex = FindId(value); }

        private void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex)
        => SelectionMade?.Invoke(control, selectedId, SelectedIndex = selectedIndex);
        #endregion
    }
}
