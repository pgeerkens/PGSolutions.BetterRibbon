﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using stdole;

using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    public abstract class AbstractSelectableModel<TSource,TCtrl> : ControlModel<TSource,TCtrl>, IControlSource
            where TSource: IControlSource
            where TCtrl: IControlVM  {
        internal AbstractSelectableModel(Func<string, IActivatable<TSource, TCtrl>> funcViewModel, IControlStrings strings)
        : base(funcViewModel,strings) { }

        #region IActivatable implementation
        public override void Detach() { SelectionMade = null;  base.Detach(); }
        #endregion

        #region IListable implementation
        public abstract IReadOnlyList<IStaticItemVM> Items { get; }

        public int FindId(string id) => Items.FindId(id);
        #endregion

        #region ISelectableList implementation
        public event SelectionMadeEventHandler SelectionMade;

        public int    SelectedIndex { get; set; }
        public string SelectedId    { get => Items[SelectedIndex].Id; set => SelectedIndex = FindId(value); }

        protected void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex)
        => SelectionMade?.Invoke(control, selectedId, SelectedIndex = selectedIndex);
        #endregion

        #region IImageable implementation
        public ImageObject Image     { get; set; } = "MacroSecurity";
        public bool        ShowImage { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;
       #endregion
    }
}
