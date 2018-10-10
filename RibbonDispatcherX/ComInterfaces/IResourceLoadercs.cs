using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>The default COM interface exposed by {ResourceLoader} objects.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IResourceLoader)]
    public interface IResourceLoader {
        /// <summary>Returns the specified {ControlStrings} object.</summary>
        [DispId(DispIds.ControlStringsIndexer)]
        [Description("Returns the specified ControlStrings object.")]
        IRibbonTextLanguageControl this[string ControlId] { get; }

        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        [DispId(DispIds.AddControlStrings)]
        [Description("Adds a new ControlString to the collection, and returns it.")]
        string AddControlStrings(string ControlId,
            [Optional]string Label,
            [Optional]string ScreenTip,
            [Optional]string SuperTip,
            [Optional]string AlternateLabel,
            [Optional]string Description,
            [Optional]string KeyTip);
    }

    internal static partial class DispIds {
        public const int ControlStringsIndexer = 1;
        public const int AddControlStrings     = 1 + ControlStringsIndexer;
    }
}
