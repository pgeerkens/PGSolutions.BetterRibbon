using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.ISelectedEvents)]
    public interface ISelectionMadeEvents {
        /// <summary>TODO</summary>
        [DispId(1)]
        void SelectionMade(string ItemId, int ItemIndex);
    }
}
