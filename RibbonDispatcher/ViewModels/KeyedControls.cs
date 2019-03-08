////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class KeyedControls: KeyedCollection<string, IControlVM> {
        internal KeyedControls() : base() { }
        internal KeyedControls(IEnumerable<IControlVM> list) : base() {
            foreach (var item in list) base.Add(item);
        }
        protected override string GetKeyForItem(IControlVM control) => control?.ControlId;

        public TCtrl Item<TCtrl>(string id) where TCtrl:IControlVM 
        => this.Contains(id) && this[id] is TCtrl ctrl ? ctrl : default;
    }
}
