using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IControlStrings)]
    public interface IControlStrings {
        /// <summary>Returns the number of items in the collection.</summary>
        int         Count         { get; }
        /// <summary>Returns the collection item at the specified index.</summary>
        string this[string Index] { get; }

        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        string AddControl(string ControlId,
            [Optional]string Label,
            [Optional]string ScreenTip,
            [Optional]string SuperTip,
            [Optional]string AlternateLabel,
            [Optional]string Description,
            [Optional]string KeyTip);
    }
}
