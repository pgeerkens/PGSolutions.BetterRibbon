////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using System.Linq;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    public static partial class Extensions {
        public static int FindId(this IReadOnlyList<IStaticItemVM> items, string id)
        => items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();

        /// <summary>Adds the specified element to the dictionary only when it is not null.</summary>
        public static void AddNotNull<TValue>(this IDictionary<string, TValue> dictionary, string itemId, TValue ctrl) {
            if (ctrl != null) { dictionary?.Add(itemId, ctrl); }
        }

        /// <summary>TODO</summary>
        public static TValue GetOrDefault<TValue>(this IReadOnlyDictionary<string, TValue> dictionary, string key) {
            if (dictionary == null) return default;
            return dictionary.TryGetValue(key ?? "", out var ctrl) ? ctrl : default;
        }
    }
}
