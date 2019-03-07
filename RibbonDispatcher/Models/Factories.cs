////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////


using PGSolutions.RibbonDispatcher.ViewModels;
using System.Collections.ObjectModel;

namespace PGSolutions.RibbonDispatcher.Models {
    public class Factories: KeyedCollection<string,ViewModelFactory> {
        protected override string GetKeyForItem(ViewModelFactory item) => item.Key;

        public bool TryGetValue(string key, out ViewModelFactory factory)
        => (factory = TryGetValue(key)) != null;

        public ViewModelFactory TryGetValue(string key) => Contains(key) ? this[key] : default;
    }
}
