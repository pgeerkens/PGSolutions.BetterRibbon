////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher2013.Utilities {
    /// <summary>TODO</summary>
    public static class DictionaryExtensions {
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
