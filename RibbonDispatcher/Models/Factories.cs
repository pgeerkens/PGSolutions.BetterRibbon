////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ViewModels;
using System.Collections.ObjectModel;

namespace PGSolutions.RibbonDispatcher.Models {
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class Factories: KeyedCollection<string,ViewModelFactory> {
        protected override string GetKeyForItem(ViewModelFactory item)
        => item?.Key ?? throw new ArgumentNullException(nameof(item));

        public bool TryGetValue(string key, out ViewModelFactory factory)
        => (factory = TryGetValue(key)) != null;

        public ViewModelFactory TryGetValue(string key) => Contains(key) ? this[key] : default;
    }
}
