using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher2013.AbstractCOM {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1711:IdentifiersShouldNotHaveIncorrectSuffix", Justification = "Necessary for COM Interop.")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IControlChangedEventArgs {
        /// <summary>The</summary>
        [DispId(DispIds.ControlId)]
        string ControlId { get; }
    }
}
