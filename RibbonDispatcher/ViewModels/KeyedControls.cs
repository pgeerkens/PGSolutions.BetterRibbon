////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.ObjectModel;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public class KeyedControls: KeyedCollection<string, IControlVM> {
        internal KeyedControls() : base() { }
        protected override string GetKeyForItem(IControlVM control) => control?.Id;

        public TCtrl Item<TCtrl>(string id) where TCtrl:IControlVM 
        => this[id] is TCtrl ctrl ? ctrl : default;
    }
}
